"""Formatos y mapeos por defecto para el generador de COAs."""

DEFAULT_MICRO_FORMATS = {
    "Estándar": {
        "tipo": "simple",
        "descripcion": "Formato estándar genérico",
        "parametros": [
            {"nombre": "Total Plate Count",             "clave": "TPC",          "defecto": "",         "ingresar": True,  "sinonimos": "total plate count,total plate,ram,tpc"},
            {"nombre": "Total Coliforms",               "clave": "Coliforms",    "defecto": "<10",      "ingresar": False, "sinonimos": "total coliforms,coliforms"},
            {"nombre": "E. coli",                       "clave": "Ecoli",        "defecto": "<10",      "ingresar": False, "sinonimos": "e. coli,e.coli"},
            {"nombre": "Yeast",                         "clave": "Yeast",        "defecto": "",         "ingresar": True,  "sinonimos": "yeast"},
            {"nombre": "Mold",                          "clave": "Mold",         "defecto": "",         "ingresar": True,  "sinonimos": "mold,moulds"},
            {"nombre": "Staphylococcus aureus",         "clave": "Staph",        "defecto": "<10",      "ingresar": False, "sinonimos": "staphylococcus,coagulase"},
            {"nombre": "Salmonella sp.",                "clave": "Salmonella",   "defecto": "Negative", "ingresar": False, "sinonimos": "salmonella"},
            {"nombre": "Listeria sp.",                  "clave": "Listeria",     "defecto": "Negative", "ingresar": False, "sinonimos": "listeria"},
        ]
    },
    "Inabata / Korea": {
        "tipo": "simple",
        "descripcion": "Inabata y clientes coreanos",
        "parametros": [
            {"nombre": "Total Plate Count",             "clave": "TPC",          "defecto": "",         "ingresar": True,  "sinonimos": "total plate count,total plate,ram"},
            {"nombre": "Total Coliforms",               "clave": "Coliforms",    "defecto": "<10",      "ingresar": False, "sinonimos": "total coliforms,coliforms"},
            {"nombre": "E. coli",                       "clave": "Ecoli",        "defecto": "Negative", "ingresar": False, "sinonimos": "e. coli,e.coli"},
            {"nombre": "Yeast",                         "clave": "Yeast",        "defecto": "",         "ingresar": True,  "sinonimos": "yeast"},
            {"nombre": "Mold",                          "clave": "Mold",         "defecto": "",         "ingresar": True,  "sinonimos": "mold,moulds"},
            {"nombre": "Staphylococcus aureus",         "clave": "Staph",        "defecto": "<10",      "ingresar": False, "sinonimos": "staphylococcus,coagulase"},
            {"nombre": "Salmonella sp.",                "clave": "Salmonella",   "defecto": "Negative", "ingresar": False, "sinonimos": "salmonella"},
            {"nombre": "Listeria sp.",                  "clave": "Listeria",     "defecto": "Negative", "ingresar": False, "sinonimos": "listeria"},
        ]
    },
    "Livemore": {
        "tipo": "simple",
        "descripcion": "Livemore / Berry Blend",
        "parametros": [
            {"nombre": "Total Plate Count",             "clave": "TPC",          "defecto": "",         "ingresar": True,  "sinonimos": "total plate count,total plate,ram"},
            {"nombre": "Total Coliforms",               "clave": "Coliforms",    "defecto": "<10",      "ingresar": False, "sinonimos": "total coliforms,coliforms"},
            {"nombre": "Enterobacteriaceae",            "clave": "Enterobacteria","defecto": "<10",      "ingresar": False, "sinonimos": "enterobacteriaceae,enterobacterias,enterobacteria"},
            {"nombre": "E. coli",                       "clave": "Ecoli",        "defecto": "<10",      "ingresar": False, "sinonimos": "e. coli,e.coli"},
            {"nombre": "Yeast",                         "clave": "Yeast",        "defecto": "",         "ingresar": True,  "sinonimos": "yeast"},
            {"nombre": "Mold",                          "clave": "Mold",         "defecto": "",         "ingresar": True,  "sinonimos": "mold,moulds"},
            {"nombre": "Staphylococcus aureus",         "clave": "Staph",        "defecto": "<10",      "ingresar": False, "sinonimos": "staphylococcus,coagulase"},
            {"nombre": "Salmonella sp.",                "clave": "Salmonella",   "defecto": "Absence",  "ingresar": False, "sinonimos": "salmonella"},
            {"nombre": "Listeria sp.",                  "clave": "Listeria",     "defecto": "Absence",  "ingresar": False, "sinonimos": "listeria"},
            {"nombre": "STEC",                          "clave": "STEC",         "defecto": "Absence",  "ingresar": False, "sinonimos": "stec"},
            {"nombre": "Hepatitis A",                   "clave": "HepA",         "defecto": "Absence",  "ingresar": False, "sinonimos": "hepatitis a,hepatitis"},
            {"nombre": "Norovirus GI",                  "clave": "NorovirusGI",  "defecto": "Absence",  "ingresar": False, "sinonimos": "norovirus gi"},
            {"nombre": "Norovirus GII",                 "clave": "NorovirusGII", "defecto": "Absence",  "ingresar": False, "sinonimos": "norovirus gii"},
        ]
    },
    "Trader Joe's": {
        "tipo": "simple",
        "descripcion": "Trader Joe's / Core Fruit",
        "parametros": [
            {"nombre": "Total Plate Count",             "clave": "TPC",          "defecto": "",         "ingresar": True,  "sinonimos": "total plate count,total plate,ram"},
            {"nombre": "Total Coliforms",               "clave": "Coliforms",    "defecto": "<10",      "ingresar": False, "sinonimos": "total coliforms,coliforms"},
            {"nombre": "E. coli",                       "clave": "Ecoli",        "defecto": "<10",      "ingresar": False, "sinonimos": "e. coli,e.coli"},
            {"nombre": "Staphylococcus aureus",         "clave": "Staph",        "defecto": "<10",      "ingresar": False, "sinonimos": "staphylococcus,coagulase"},
            {"nombre": "Salmonella sp.",                "clave": "Salmonella",   "defecto": "Negative", "ingresar": False, "sinonimos": "salmonella"},
            {"nombre": "Listeria sp.",                  "clave": "Listeria",     "defecto": "Negative", "ingresar": False, "sinonimos": "listeria"},
            {"nombre": "E. Coli O157",                  "clave": "EcoliO157",    "defecto": "Negative", "ingresar": False, "sinonimos": "e. coli o157,ecoli o157"},
        ]
    },
    "VLM Kit Estándar": {
        "tipo": "simple",
        "descripcion": "VLM James Farm, Lidl, Kit estándar",
        "parametros": [
            {"nombre": "Total Plate Count",             "clave": "TPC",          "defecto": "",         "ingresar": True,  "sinonimos": "total plate count,total plate,ram"},
            {"nombre": "Total Coliforms",               "clave": "Coliforms",    "defecto": "<10",      "ingresar": False, "sinonimos": "total coliforms,coliforms"},
            {"nombre": "E. coli",                       "clave": "Ecoli",        "defecto": "<10",      "ingresar": False, "sinonimos": "e. coli,e.coli"},
            {"nombre": "Yeast",                         "clave": "Yeast",        "defecto": "",         "ingresar": True,  "sinonimos": "yeast"},
            {"nombre": "Mold",                          "clave": "Mold",         "defecto": "",         "ingresar": True,  "sinonimos": "mold,moulds"},
            {"nombre": "Staphylococcus aureus",         "clave": "Staph",        "defecto": "<10",      "ingresar": False, "sinonimos": "staphylococcus,coagulase"},
            {"nombre": "Salmonella sp.",                "clave": "Salmonella",   "defecto": "Negative", "ingresar": False, "sinonimos": "salmonella"},
            {"nombre": "Listeria sp.",                  "clave": "Listeria",     "defecto": "Negative", "ingresar": False, "sinonimos": "listeria"},
            {"nombre": "Hepatitis A",                   "clave": "HepA",         "defecto": "Negative", "ingresar": False, "sinonimos": "hepatitis a,hepatitis"},
            {"nombre": "Norovirus",                     "clave": "Norovirus",    "defecto": "Negative", "ingresar": False, "sinonimos": "norovirus"},
        ]
    },
    "Woodland Partners": {
        "tipo": "simple",
        "descripcion": "Woodland Partners convencional y orgánico",
        "parametros": [
            {"nombre": "Total Plate Count",             "clave": "TPC",          "defecto": "",         "ingresar": True,  "sinonimos": "total plate count,total plate,ram"},
            {"nombre": "Total Coliforms",               "clave": "Coliforms",    "defecto": "<10",      "ingresar": False, "sinonimos": "total coliforms,coliforms"},
            {"nombre": "E. coli",                       "clave": "Ecoli",        "defecto": "<10",      "ingresar": False, "sinonimos": "e. coli,e.coli"},
            {"nombre": "Yeast",                         "clave": "Yeast",        "defecto": "",         "ingresar": True,  "sinonimos": "yeast"},
            {"nombre": "Mold",                          "clave": "Mold",         "defecto": "",         "ingresar": True,  "sinonimos": "mold,moulds"},
            {"nombre": "Staphylococcus aureus",         "clave": "Staph",        "defecto": "<10",      "ingresar": False, "sinonimos": "staphylococcus,coagulase"},
            {"nombre": "Salmonella sp.",                "clave": "Salmonella",   "defecto": "Negative", "ingresar": False, "sinonimos": "salmonella"},
            {"nombre": "Listeria sp.",                  "clave": "Listeria",     "defecto": "Negative", "ingresar": False, "sinonimos": "listeria"},
        ]
    },
    "Woolworths / Macro": {
        "tipo": "simple",
        "descripcion": "Woolworths y Macro",
        "parametros": [
            {"nombre": "Total Plate Count",                    "clave": "TPC",          "defecto": "",         "ingresar": True,  "sinonimos": "total plate count,total plate,ram"},
            {"nombre": "Total Coliforms",                      "clave": "Coliforms",    "defecto": "<10",      "ingresar": False, "sinonimos": "total coliforms,coliforms"},
            {"nombre": "E. coli",                              "clave": "Ecoli",        "defecto": "<10",      "ingresar": False, "sinonimos": "e. coli,e.coli"},
            {"nombre": "Yeast",                                "clave": "Yeast",        "defecto": "",         "ingresar": True,  "sinonimos": "yeast"},
            {"nombre": "Mold",                                 "clave": "Mold",         "defecto": "",         "ingresar": True,  "sinonimos": "mold,moulds"},
            {"nombre": "Coagulase-positive staphylococci",     "clave": "Staph",        "defecto": "<10",      "ingresar": False, "sinonimos": "coagulase-positive staphylococci,staphylococcus,coagulase"},
            {"nombre": "Salmonella sp.",                       "clave": "Salmonella",   "defecto": "Negative", "ingresar": False, "sinonimos": "salmonella"},
            {"nombre": "Listeria sp.",                         "clave": "Listeria",     "defecto": "Negative", "ingresar": False, "sinonimos": "listeria"},
            {"nombre": "Hepatitis A",                          "clave": "HepA",         "defecto": "Negative", "ingresar": False, "sinonimos": "hepatitis a,hepatitis"},
            {"nombre": "Norovirus GI",                         "clave": "NorovirusGI",  "defecto": "Absence",  "ingresar": False, "sinonimos": "norovirus gi"},
            {"nombre": "Norovirus GII",                        "clave": "NorovirusGII", "defecto": "Absence",  "ingresar": False, "sinonimos": "norovirus gii"},
        ]
    },
    "Camerican": {
        "tipo": "camerican",
        "descripcion": "Camerican — tabla por lote con Lot: en encabezado",
        "parametros": [
            {"nombre": "Total Plate Count (T.V.C.)", "clave": "TPC",        "defecto": "",         "ingresar": True,  "sinonimos": "total plate count,total plate,ram,t.v.c."},
            {"nombre": "Total Coliforms",            "clave": "Coliforms",  "defecto": "<10",      "ingresar": False, "sinonimos": "total coliforms,coliforms"},
            {"nombre": "E. coli",                    "clave": "Ecoli",      "defecto": "<10",      "ingresar": False, "sinonimos": "e. coli,e.coli"},
            {"nombre": "Yeast and mold",             "clave": "YeastMold",  "defecto": "",         "ingresar": True,  "sinonimos": "yeast and mold,yeast & mold"},
            {"nombre": "Staphylococcus aureus",      "clave": "Staph",      "defecto": "<10",      "ingresar": False, "sinonimos": "staphylococcus,coagulase"},
            {"nombre": "Salmonella",                 "clave": "Salmonella", "defecto": "Negative", "ingresar": False, "sinonimos": "salmonella"},
            {"nombre": "Listeria Monocytogenes",     "clave": "Listeria",   "defecto": "Negative", "ingresar": False, "sinonimos": "listeria"},
        ]
    },
    "VLM Great Value": {
        "tipo": "n_series",
        "descripcion": "VLM Great Value — n1 a n5 por parámetro por lote",
        "parametros": [
            {"nombre": "Total Plate Count", "clave": "TPC",      "defecto": "",    "ingresar": True,  "n_series": True,  "sinonimos": "total plate count,total plate,ram"},
            {"nombre": "Total Coliforms",   "clave": "Coliforms","defecto": "<10", "ingresar": False, "n_series": True,  "sinonimos": "total coliforms,coliforms"},
            {"nombre": "E. coli",           "clave": "Ecoli",    "defecto": "<10", "ingresar": False, "n_series": True,  "sinonimos": "e. coli,e.coli"},
            {"nombre": "Moulds",            "clave": "Mold",     "defecto": "",    "ingresar": True,  "n_series": True,  "sinonimos": "moulds,mold"},
            {"nombre": "Yeasts",            "clave": "Yeast",    "defecto": "",    "ingresar": True,  "n_series": True,  "sinonimos": "yeasts,yeast"},
            {"nombre": "Listeria Monocytogenes", "clave": "Listeria",  "defecto": "None detected", "ingresar": False, "n_series": False, "sinonimos": "listeria"},
            {"nombre": "Salmonella",             "clave": "Salmonella","defecto": "None detected", "ingresar": False, "n_series": False, "sinonimos": "salmonella"},
        ]
    },
}

DEFAULT_QUALITY_FORMATS = {
    "Mixes / Berries": {
        "descripcion": "Mix de frutas y berries en general",
        "columnas": 2,
        "parametros": [
            {"nombre": "Stem",                    "defecto": "0.00"},
            {"nombre": "Decay",                   "defecto": "0.00"},
            {"nombre": "Insect, mold, sun Damage", "defecto": ""},
            {"nombre": "Color variation",          "defecto": ""},
            {"nombre": "Overmatured / Crushed",    "defecto": ""},
            {"nombre": "Splits / Crumble",         "defecto": ""},
            {"nombre": "Vegetable Matter",         "defecto": "0.00"},
        ]
    },
    "Mora": {
        "descripcion": "Mora (Blackberry)",
        "columnas": 2,
        "parametros": [
            {"nombre": "Stem",                    "defecto": "0.00"},
            {"nombre": "Decay",                   "defecto": "0.00"},
            {"nombre": "Insect, mold, sun Damage", "defecto": ""},
            {"nombre": "Color variation",          "defecto": ""},
            {"nombre": "Overmatured / Crushed",    "defecto": ""},
            {"nombre": "Splits",                   "defecto": ""},
            {"nombre": "Vegetable Matter",         "defecto": "0.00"},
        ]
    },
    "Arándano": {
        "descripcion": "Arándano (Blueberry)",
        "columnas": 2,
        "parametros": [
            {"nombre": "Decay",                   "defecto": "0.00"},
            {"nombre": "Insect, mould, sun damage","defecto": "0.00"},
            {"nombre": "Russet",                   "defecto": ""},
            {"nombre": "Dehydration",              "defecto": ""},
            {"nombre": "Color Variation",          "defecto": ""},
            {"nombre": "Overmatured / crushed",    "defecto": ""},
            {"nombre": "Vegetable Matter",         "defecto": "0.00"},
            {"nombre": "Foreign Matter",           "defecto": "0.00"},
        ]
    },
    "Cereza": {
        "descripcion": "Cereza (Cherry)",
        "columnas": 2,
        "parametros": [
            {"nombre": "Color Variation",          "defecto": ""},
            {"nombre": "Pits",                     "defecto": "0.00"},
            {"nombre": "Stem / Stalks",            "defecto": "0.00"},
            {"nombre": "Mould, sun damage",        "defecto": ""},
            {"nombre": "Insect damage",            "defecto": ""},
            {"nombre": "Overmatured / Crushed",    "defecto": "0.00"},
            {"nombre": "Broken / Crumble",         "defecto": ""},
            {"nombre": "Blocked",                  "defecto": "0.00"},
            {"nombre": "Foreign Material",         "defecto": "0.00"},
            {"nombre": "Vegetable Material",       "defecto": "0.50"},
        ]
    },
    "Frutilla": {
        "descripcion": "Frutilla (Strawberry)",
        "columnas": 2,
        "parametros": [
            {"nombre": "Color Variation",          "defecto": ""},
            {"nombre": "Stem / Stalks",            "defecto": ""},
            {"nombre": "Mould, sun damage",        "defecto": ""},
            {"nombre": "Insect damage",            "defecto": ""},
            {"nombre": "Overmatured / Crushed",    "defecto": "0.00"},
            {"nombre": "Broken / Crumble",         "defecto": ""},
            {"nombre": "Blocked",                  "defecto": "0.00"},
            {"nombre": "Foreign Material",         "defecto": "0.00"},
            {"nombre": "Vegetable Material",       "defecto": "0.50"},
        ]
    },
    "Frambuesa / Crumble": {
        "descripcion": "Frambuesa y crumble (3 columnas: Limit + Results)",
        "columnas": 3,
        "parametros": [
            {"nombre": "Stem",                    "limit": "1 unit",    "defecto": ""},
            {"nombre": "Foreign Vegetable Matter", "limit": "4 units",   "defecto": ""},
            {"nombre": "Foreign Matter",           "limit": "None",      "defecto": "None"},
            {"nombre": "Insect Count",             "limit": "2 per box", "defecto": ""},
            {"nombre": "Larvae Count",             "limit": "8 per kg",  "defecto": ""},
            {"nombre": "Whole and broken fruit",   "limit": "<10%",      "defecto": ""},
        ]
    },
    "Frambuesa": {
        "descripcion": "Frambuesa (2 columnas)",
        "columnas": 2,
        "parametros": [
            {"nombre": "Stem",                     "defecto": "0.00"},
            {"nombre": "Decay",                    "defecto": "0.00"},
            {"nombre": "Insect, mold, sun Damage", "defecto": "0.00"},
            {"nombre": "Color variation",          "defecto": "0.00"},
            {"nombre": "Overmatured / Crushed",    "defecto": ""},
            {"nombre": "Splits / Crumble",         "defecto": ""},
            {"nombre": "Vegetable Matter",         "defecto": "0.00"},
        ]
    },
    "Arilos": {
        "descripcion": "Arilos (granada)",
        "columnas": 2,
        "parametros": [
            {"nombre": "Mold, Sun Damage",         "defecto": "0.00"},
            {"nombre": "Vegetal matter from fruits","defecto": "0.00"},
        ]
    },
    "Palta": {
        "descripcion": "Palta (Avocado)",
        "columnas": 2,
        "parametros": [
            {"nombre": "Color Variation",          "defecto": "0.00"},
            {"nombre": "Stem",                     "defecto": "0.00"},
            {"nombre": "Pits",                     "defecto": ""},
            {"nombre": "Mould, sun damage",        "defecto": "0.00"},
            {"nombre": "Overmatured and crushed",  "defecto": ""},
            {"nombre": "Broken and small pieces",  "defecto": ""},
            {"nombre": "Foreign Material",         "defecto": "0.00"},
            {"nombre": "Vegetal from fruit",       "defecto": "0.00"},
            {"nombre": "Insect Count",             "defecto": "0.00"},
        ]
    },
    "Banana": {
        "descripcion": "Banana",
        "columnas": 2,
        "parametros": [
            {"nombre": "Stem",                     "defecto": "0.00"},
            {"nombre": "Decay",                    "defecto": "0.00"},
            {"nombre": "Insect, mold, sun Damage", "defecto": "0.00"},
            {"nombre": "Color variation",          "defecto": ""},
            {"nombre": "Overmatured / Crushed",    "defecto": ""},
            {"nombre": "Splits",                   "defecto": ""},
            {"nombre": "Vegetable Matter",         "defecto": "0.00"},
        ]
    },
    "Dragon Fruit": {
        "descripcion": "Pitahaya / Dragon Fruit",
        "columnas": 2,
        "parametros": [
            {"nombre": "Decay",                    "defecto": "0.00"},
            {"nombre": "Insect, mould, sun damage","defecto": "0.00"},
            {"nombre": "Color Variation",          "defecto": ""},
            {"nombre": "Overmatured / Crushed",    "defecto": ""},
            {"nombre": "Broken / Small Pieces",    "defecto": ""},
            {"nombre": "Blocked",                  "defecto": ""},
            {"nombre": "Vegetable Matter",         "defecto": "0.00"},
            {"nombre": "Foreign Matter",           "defecto": "0.00"},
        ]
    },
    "Mango": {
        "descripcion": "Mango",
        "columnas": 2,
        "parametros": [
            {"nombre": "Stem",                     "defecto": "0.00"},
            {"nombre": "Decay",                    "defecto": "0.00"},
            {"nombre": "Insect, mold, sun Damage", "defecto": "0.00"},
            {"nombre": "Color variation",          "defecto": ""},
            {"nombre": "Overmatured / Crushed",    "defecto": ""},
            {"nombre": "Blocked",                  "defecto": ""},
            {"nombre": "Vegetable Matter",         "defecto": "0.00"},
        ]
    },
    "Durazno": {
        "descripcion": "Durazno (Peach)",
        "columnas": 2,
        "parametros": [
            {"nombre": "Stem",                     "defecto": "0.00"},
            {"nombre": "Decay",                    "defecto": "0.00"},
            {"nombre": "Insect, mold, sun Damage", "defecto": "0.00"},
            {"nombre": "Color variation",          "defecto": "0.00"},
            {"nombre": "Overmatured / Crushed",    "defecto": ""},
            {"nombre": "Splits",                   "defecto": ""},
            {"nombre": "Vegetable Matter",         "defecto": "0.00"},
        ]
    },
    "Piña": {
        "descripcion": "Piña (Pineapple)",
        "columnas": 2,
        "parametros": [
            {"nombre": "Decay",                        "defecto": ""},
            {"nombre": "Mechanical, insect, mold, sun damage", "defecto": ""},
            {"nombre": "Color variation",              "defecto": ""},
            {"nombre": "Overmatured / Crushed",        "defecto": ""},
            {"nombre": "Small pieces / Splits",        "defecto": ""},
            {"nombre": "Vegetable Matter",             "defecto": ""},
            {"nombre": "Foreign Matter",               "defecto": ""},
        ]
    },
}

DEFAULT_CLIENT_MAP = [
    {"keyword": "camerican",   "formato": "Camerican"},
    {"keyword": "livemore",    "formato": "Livemore"},
    {"keyword": "trader joe",  "formato": "Trader Joe's"},
    {"keyword": "inabata",     "formato": "Inabata / Korea"},
    {"keyword": "korea",       "formato": "Inabata / Korea"},
    {"keyword": "woodland",    "formato": "Woodland Partners"},
    {"keyword": "woolworths",  "formato": "Woolworths / Macro"},
    {"keyword": "macro",       "formato": "Woolworths / Macro"},
    {"keyword": "vlm",         "formato": "VLM Kit Estandar"},
    {"keyword": "great value", "formato": "VLM Great Value"},
]
DEFAULT_BRAND_MAP = [
    {"keyword": "food lion",        "formato": "Woodland Partners"},
    {"keyword": "bowl & basket",    "formato": "Woodland Partners"},
    {"keyword": "bowl and basket",  "formato": "Woodland Partners"},
    {"keyword": "nature's promise", "formato": "Woodland Partners"},
    {"keyword": "natures promise",  "formato": "Woodland Partners"},
    {"keyword": "wholesome pantry", "formato": "Woodland Partners"},
    {"keyword": "shoprite",         "formato": "Woodland Partners"},
    {"keyword": "hannaford",        "formato": "Woodland Partners"},
]
