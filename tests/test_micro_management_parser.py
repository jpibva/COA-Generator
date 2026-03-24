from micro_management_app import _extract_micro_from_text


def _first_result(raw_text):
    items = _extract_micro_from_text(raw_text, source_pdf="x.pdf")
    assert items
    return items[0]["resultados"]


def test_extract_supports_parameter_and_value_on_separate_lines():
    raw = """
    Product: Organic 5 Fruit Blend
    Lot: 60351
    Total Plate Count
    10^2
    E. coli
    <10
    Salmonella
    Negative
    """
    res = _first_result(raw)
    assert res["TPC"] == "100"
    assert res["Ecoli"] == "<10"
    assert res["Salmonella"].lower() == "negative"


def test_extract_supports_n_series_split_lines():
    raw = """
    Producto: VLM Mix
    Lote: 70001
    Total Plate Count (1)
    10^3
    Total Plate Count (2)
    10^2
    """
    res = _first_result(raw)
    assert res["TPC_n1"] == "1000"
    assert res["TPC_n2"] == "100"
