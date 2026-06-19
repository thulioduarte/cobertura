import importlib
mods = ["streamlit", "pandas", "numpy", "openpyxl", "xlsxwriter"]
for m in mods:
    importlib.import_module(m)
print("Instalação OK. Bibliotecas principais carregadas.")
