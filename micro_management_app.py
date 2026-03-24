import os
import re
import unicodedata
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from datetime import datetime

from coa_storage import (
    CONFIG_FILE,
    MICRO_AUDIT_FILE,
    load_config,
    load_micro_history,
    normalize_product_name,
    save_micro_history_record,
    append_micro_audit_rows,
)
from coa_formats import DEFAULT_CLIENT_MAP, DEFAULT_BRAND_MAP

config = load_config(CONFIG_FILE)


def _normalize_text_token(text):
    text = unicodedata.normalize("NFKD", text or "")
    text = "".join(ch for ch in text if not unicodedata.combining(ch))
    return re.sub(r"[^a-z0-9]+", " ", text.lower()).strip()


def _canonical_product_name(name):
    aliases = config.get("product_alias_map", {}) or {}
    norm = _normalize_text_token(name)
    for canonical, alias_list in aliases.items():
        if _normalize_text_token(canonical) == norm:
            return canonical
        for alias in alias_list or []:
            if _normalize_text_token(alias) == norm:
                return canonical
    return name


def _detect_micro_format(cliente, producto=""):
    cliente_low = (cliente or "").lower()
    producto_low = (producto or "").lower()

    def _resolve(fmt_name):
        for cfg_name in config.get("micro_formats", {}):
            if cfg_name.lower().replace(" ", "").startswith(fmt_name.lower().replace(" ", "")[:6]):
                return cfg_name
        return fmt_name

    for entry in config.get("client_map", DEFAULT_CLIENT_MAP):
        if entry["keyword"].lower() in cliente_low:
            return _resolve(entry["formato"])

    for entry in config.get("brand_map", DEFAULT_BRAND_MAP):
        if entry["keyword"].lower() in producto_low:
            return _resolve(entry["formato"])

    keys = list(config.get("micro_formats", {}).keys())
    return keys[0] if keys else ""


def _normalize_micro_value(raw_value):
    if raw_value is None:
        return ""
    value = str(raw_value).strip()
    if not value:
        return ""

    v = value.replace("×", "x").replace("X", "x")
    if re.fullmatch(r"\d+\^\d+", v):
        b, e = v.split("^")
        try:
            return str(int(float(b) ** int(e)))
        except Exception:
            return value

    sci = re.match(r"^\s*([0-9]+(?:[.,][0-9]+)?)\s*(?:x\s*10\^?|e)\s*([+-]?\d+)\s*$", v, re.IGNORECASE)
    if sci:
        base = sci.group(1).replace(",", ".")
        exp = sci.group(2)
        try:
            return str(int(round(float(base) * (10 ** int(exp)))))
        except Exception:
            return value
    return value


def _extract_text_from_pdf(pdf_path):
    text = ""
    try:
        import pdfplumber
        with pdfplumber.open(pdf_path) as pdf:
            text = "\n".join((p.extract_text() or "") for p in pdf.pages)
    except Exception:
        text = ""
    if text.strip():
        return text

    # OCR fallback opcional
    try:
        from pdf2image import convert_from_path
        import pytesseract
        imgs = convert_from_path(pdf_path, dpi=300)
        return "\n".join(pytesseract.image_to_string(i, lang="eng+spa") for i in imgs)
    except Exception:
        return ""


def _extract_micro_from_text(raw_text, source_pdf=""):
    if not raw_text:
        return []

    lines = [ln.strip() for ln in raw_text.replace("\r", "\n").splitlines() if ln.strip()]
    lot_re = re.compile(r"\b(\d{5}(?:-\d+)?)\b")
    product_re = re.compile(r"(?:product|producto|sample|muestra)\s*[:\-]\s*(.+)$", re.IGNORECASE)

    metric_patterns = {
        "TPC": re.compile(r"(total\s*plate|t\.?p\.?c|ram|recuento\s*aerobios?)", re.IGNORECASE),
        "Coliforms": re.compile(r"(total\s*coliform|coliformes?)", re.IGNORECASE),
        "Enterobacteria": re.compile(r"(enterobacter)", re.IGNORECASE),
        "Ecoli": re.compile(r"(e\.?\s*coli)", re.IGNORECASE),
        "Yeast": re.compile(r"(yeast|levadur)", re.IGNORECASE),
        "Mold": re.compile(r"(mold|mould|hong)", re.IGNORECASE),
        "Staph": re.compile(r"(staph|staphylococcus|coagulase)", re.IGNORECASE),
        "Salmonella": re.compile(r"(salmonella)", re.IGNORECASE),
        "Listeria": re.compile(r"(listeria)", re.IGNORECASE),
    }
    value_re = re.compile(
        r"(?:(?:<|>)\s*\d+(?:[.,]\d+)?)|(?:\d+\^\d+)|(?:\d+(?:[.,]\d+)?\s*(?:x|×)\s*10\^?\s*\d+)|(?:\d+(?:[.,]\d+)?e[+-]?\d+)|(?:\d+(?:[.,]\d+)?)|(?:negative|negativo|absence|absent|not\s*detected)",
        re.IGNORECASE,
    )

    results = []
    current = {"producto": "", "lote": "", "resultados": {}, "source_pdf": os.path.basename(source_pdf or "")}

    def flush():
        if current["lote"] and current["resultados"]:
            results.append({
                "producto": _canonical_product_name(current["producto"] or "Producto sin nombre"),
                "producto_original": current["producto"] or "",
                "lote": current["lote"],
                "resultados": dict(current["resultados"]),
                "source_pdf": current["source_pdf"],
            })

    pending_params = []  # parámetros detectados sin valor aún

    def _capture_value_for_pending(line):
        if not pending_params:
            return False
        vm = value_re.search(line)
        if not vm:
            return False
        val = _normalize_micro_value(vm.group(0))
        used = False
        for pkey, nidx in list(pending_params):
            if nidx:
                current["resultados"][f"{pkey}_n{nidx}"] = val
            else:
                current["resultados"][pkey] = val
            used = True
        pending_params.clear()
        return used

    for ln in lines:
        # Si la línea trae solo valor (y quedó parámetro pendiente en línea previa), completar.
        if _capture_value_for_pending(ln):
            continue

        pm = product_re.search(ln)
        if pm:
            current["producto"] = pm.group(1).strip()

        lots = lot_re.findall(ln)
        if lots:
            if current["lote"] and current["resultados"] and lots[0] != current["lote"]:
                flush()
                current["resultados"] = {}
            current["lote"] = lots[0].split("-")[0]

        for key, patt in metric_patterns.items():
            if not patt.search(ln):
                continue
            vm = value_re.search(ln)
            n_match = re.search(r"\((\d)\)", ln)
            n_idx = None
            if n_match:
                idx = int(n_match.group(1))
                if 1 <= idx <= 5 and key in {"TPC", "Coliforms", "Ecoli", "Yeast", "Mold"}:
                    n_idx = idx

            if vm and n_idx and vm.group(0).strip() == str(n_idx) and f"({n_idx})" in ln:
                vm = None

            if vm:
                val = _normalize_micro_value(vm.group(0))
                if n_idx:
                    current["resultados"][f"{key}_n{n_idx}"] = val
                else:
                    current["resultados"][key] = val
            else:
                # parámetro detectado pero sin valor en la misma línea
                pending_params.append((key, n_idx))

    flush()
    return results


def _extract_micro_from_pdf(pdf_path):
    return _extract_micro_from_text(_extract_text_from_pdf(pdf_path), source_pdf=pdf_path)


def _merge_blocks(blocks):
    grouped = {}
    sources = {}
    for b in blocks:
        prod = _canonical_product_name(b.get("producto", "") or "Producto sin nombre")
        lote = (b.get("lote", "") or "").split("-")[0]
        if not lote:
            continue
        key = (prod, lote)
        grouped.setdefault(key, {}).update(b.get("resultados", {}))
        sources.setdefault(key, set())
        if b.get("source_pdf"):
            sources[key].add(b["source_pdf"])

    productos_lotes, formatos, prefill, meta = {}, {}, {}, {}
    for (prod, lote), vals in grouped.items():
        productos_lotes.setdefault(prod, []).append(lote)
        prefill.setdefault(prod, {})[lote] = vals
        formatos.setdefault(prod, _detect_micro_format("", prod))
        meta[(prod, lote)] = {"sources": sorted(list(sources.get((prod, lote), set())))}

    for p in productos_lotes:
        productos_lotes[p] = sorted(list(dict.fromkeys(productos_lotes[p])))
    return productos_lotes, formatos, prefill, meta


class MicroReviewDialog(tk.Toplevel):
    def __init__(self, parent, cliente, productos_lotes, formatos, prefill, history):
        super().__init__(parent)
        self.title("Revisión Microbiología")
        self.geometry("980x700")
        self.grab_set()
        self.cliente = cliente
        self.productos_lotes = productos_lotes
        self.formatos = formatos
        self.prefill = prefill
        self.history = history
        self.vars = {}
        self.format_vars = {}
        self.confirmed = False
        self._build_ui()

    def _build_ui(self):
        nb = ttk.Notebook(self)
        nb.pack(fill=tk.BOTH, expand=True, padx=8, pady=8)

        for producto, lotes in self.productos_lotes.items():
            tab = ttk.Frame(nb)
            nb.add(tab, text=producto[:28])
            self._build_product_tab(tab, producto, lotes)

        btn = ttk.Frame(self)
        btn.pack(fill=tk.X, padx=8, pady=(0, 8))
        ttk.Button(btn, text="Cancelar", command=self.destroy).pack(side=tk.RIGHT, padx=4)
        ttk.Button(btn, text="Guardar", command=self._on_save).pack(side=tk.RIGHT)

    def _build_product_tab(self, parent, producto, lotes):
        top = ttk.Frame(parent)
        top.pack(fill=tk.X, padx=8, pady=(8, 4))
        ttk.Label(top, text="Formato:").pack(side=tk.LEFT)
        fmt_var = tk.StringVar(value=self.formatos.get(producto, _detect_micro_format(self.cliente, producto)))
        self.format_vars[producto] = fmt_var
        cmb = ttk.Combobox(top, textvariable=fmt_var, values=list(config.get("micro_formats", {}).keys()), state="readonly", width=28)
        cmb.pack(side=tk.LEFT, padx=6)

        notebook_lotes = ttk.Notebook(parent)
        notebook_lotes.pack(fill=tk.BOTH, expand=True, padx=8, pady=8)

        self.vars[producto] = {}
        for lote in lotes:
            tab = ttk.Frame(notebook_lotes)
            notebook_lotes.add(tab, text=f"Lote {lote}")
            self.vars[producto][lote] = {}

            fmt = config.get("micro_formats", {}).get(fmt_var.get(), {})
            params = fmt.get("parametros", [])
            tipo = fmt.get("tipo", "simple")

            hkey = (self.cliente.lower().strip(), normalize_product_name(producto), lote)
            hist = self.history.get(hkey, {})
            pref = self.prefill.get(producto, {}).get(lote, {})

            row = 0
            for p in params:
                clave = p["clave"]
                nombre = p["nombre"]
                defecto = p.get("defecto", "")
                n_series = p.get("n_series", False)

                if tipo == "n_series" and n_series:
                    ttk.Label(tab, text=nombre, font=("Segoe UI", 9, "bold")).grid(row=row, column=0, sticky="w", padx=6, pady=(8, 2))
                    row += 1
                    for n in range(1, 6):
                        k = f"{clave}_n{n}"
                        ttk.Label(tab, text=f"n{n}").grid(row=row, column=(n-1)*2, sticky="e", padx=(6, 2), pady=2)
                        v = tk.StringVar(value=str(pref.get(k, hist.get(k, defecto) or "")))
                        ttk.Entry(tab, textvariable=v, width=10).grid(row=row, column=(n-1)*2+1, sticky="w", padx=(0, 8), pady=2)
                        self.vars[producto][lote][k] = v
                    row += 1
                else:
                    ttk.Label(tab, text=nombre, width=28, anchor="w").grid(row=row, column=0, sticky="w", padx=6, pady=3)
                    v = tk.StringVar(value=str(pref.get(clave, hist.get(clave, defecto) or "")))
                    ttk.Entry(tab, textvariable=v, width=20).grid(row=row, column=1, sticky="w", padx=6, pady=3)
                    self.vars[producto][lote][clave] = v
                    row += 1

    def _on_save(self):
        self.confirmed = True
        self.destroy()

    def results(self):
        out = {}
        for p, lotes in self.vars.items():
            out[p] = {}
            for l, fields in lotes.items():
                out[p][l] = {k: v.get().strip() for k, v in fields.items()}
        return out

    def formatos_finales(self):
        return {p: v.get() for p, v in self.format_vars.items()}


class MicroManagementApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Módulo de Gestión Microbiológica")
        self.geometry("920x560")
        self.pdf_paths = []
        self._build_ui()

    def _build_ui(self):
        frm = ttk.Frame(self)
        frm.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        self.cliente = tk.StringVar()
        self.producto = tk.StringVar()
        self.lotes = tk.StringVar()

        for i, (lab, var) in enumerate([
            ("Cliente", self.cliente),
            ("Producto (manual)", self.producto),
            ("Lotes (coma)", self.lotes),
        ]):
            row = ttk.Frame(frm)
            row.pack(fill=tk.X, pady=4)
            ttk.Label(row, text=lab + ":", width=20).pack(side=tk.LEFT)
            ttk.Entry(row, textvariable=var).pack(side=tk.LEFT, fill=tk.X, expand=True)

        brow = ttk.Frame(frm)
        brow.pack(fill=tk.X, pady=6)
        ttk.Button(brow, text="Seleccionar PDF(s)", command=self._pick_pdfs).pack(side=tk.LEFT)
        ttk.Button(brow, text="Limpiar", command=self._clear_pdfs).pack(side=tk.LEFT, padx=6)
        self.lbl_pdfs = ttk.Label(brow, text="Sin PDF")
        self.lbl_pdfs.pack(side=tk.LEFT, padx=10)

        ttk.Button(frm, text="Procesar y revisar", command=self._run_flow).pack(anchor="w", pady=8)

        self.log = tk.Text(frm, height=16)
        self.log.pack(fill=tk.BOTH, expand=True)

    def _pick_pdfs(self):
        paths = filedialog.askopenfilenames(
            title="Selecciona PDF(s) de microbiología",
            filetypes=[("PDF files", "*.pdf")],
            initialdir=config.get("pdf_folder", "")
        )
        if paths:
            self.pdf_paths = list(paths)
            self.lbl_pdfs.configure(text=f"{len(self.pdf_paths)} PDF(s) seleccionados")

    def _clear_pdfs(self):
        self.pdf_paths = []
        self.lbl_pdfs.configure(text="Sin PDF")

    def _run_flow(self):
        cliente = self.cliente.get().strip()
        producto_manual = self.producto.get().strip()
        lotes_manual = [x.strip() for x in self.lotes.get().split(",") if x.strip()]

        blocks = []
        for p in self.pdf_paths:
            ext = _extract_micro_from_pdf(p)
            blocks.extend(ext)
            self.log.insert(tk.END, f"PDF {os.path.basename(p)}: {len(ext)} bloque(s) detectado(s)\n")

        productos_lotes, formatos, prefill, meta = _merge_blocks(blocks) if blocks else ({}, {}, {}, {})

        if producto_manual and lotes_manual:
            canon = _canonical_product_name(producto_manual)
            productos_lotes.setdefault(canon, [])
            productos_lotes[canon].extend(lotes_manual)
            productos_lotes[canon] = sorted(list(dict.fromkeys(productos_lotes[canon])))
            formatos.setdefault(canon, _detect_micro_format(cliente, canon))
            prefill.setdefault(canon, {})
            for lt in lotes_manual:
                prefill[canon].setdefault(lt, {})

        if not productos_lotes:
            messagebox.showwarning("Sin datos", "Carga PDF(s) o completa Producto + Lotes.")
            return

        history = load_micro_history()
        dlg = MicroReviewDialog(self, cliente, productos_lotes, formatos, prefill, history)
        self.wait_window(dlg)
        if not dlg.confirmed:
            return

        results = dlg.results()
        formatos_fin = dlg.formatos_finales()
        audit_rows = []

        for prod, lotes_map in results.items():
            fmt = formatos_fin.get(prod, _detect_micro_format(cliente, prod))
            for lote, vals in lotes_map.items():
                hkey = (cliente.lower().strip(), normalize_product_name(prod), lote)
                hist = history.get(hkey, {})
                final_values = dict(hist)
                for k, v in vals.items():
                    if v:
                        final_values[k] = _normalize_micro_value(v)

                diffs = []
                for k, nv in final_values.items():
                    ov = str(hist.get(k, "")).strip()
                    nv = str(nv).strip()
                    if k not in {"Cliente", "Producto", "Lote", "Fecha"} and ov != nv:
                        diffs.append((k, ov, nv))

                if hist and diffs:
                    preview = "\n".join(f"• {k}: '{o}' -> '{n}'" for k, o, n in diffs[:8])
                    if not messagebox.askyesno("Remuestreo detectado", f"{prod} / Lote {lote}\n\n{preview}\n\n¿Aplicar cambios?"):
                        continue

                save_micro_history_record(cliente, prod, lote, final_values, config, formato_nombre=fmt)

                src = ", ".join(meta.get((prod, lote), {}).get("sources", []))
                for k, ov, nv in diffs:
                    audit_rows.append({
                        "FechaHora": datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
                        "Cliente": cliente,
                        "ProductoOriginal": prod,
                        "ProductoCanonico": _canonical_product_name(prod),
                        "Lote": lote,
                        "Formato": fmt,
                        "Parametro": k,
                        "ValorAnterior": ov,
                        "ValorNuevo": nv,
                        "OrigenPDF": src,
                        "Motivo": "resample_merge",
                    })

        if audit_rows:
            audit_file = config.get("micro_module", {}).get("audit_file", MICRO_AUDIT_FILE)
            append_micro_audit_rows(audit_rows, audit_file=audit_file)

        messagebox.showinfo("Listo", "Gestión microbiológica finalizada.")


if __name__ == "__main__":
    app = MicroManagementApp()
    app.mainloop()
