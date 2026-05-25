# -*- coding: utf-8 -*-
from pathlib import Path
import importlib.util
import sys

BASE = Path(__file__).resolve().parent
ENGINE = BASE / "gerar_estudo_cobertura_anexo_corrigido.py"

print("Pasta do app:", BASE)
print("Motor encontrado:", ENGINE.exists(), "-", ENGINE)

if not ENGINE.exists():
    raise SystemExit("ERRO: motor não encontrado na mesma pasta do app.py")

spec = importlib.util.spec_from_file_location("gerar_estudo_cobertura_anexo_corrigido", ENGINE)
mod = importlib.util.module_from_spec(spec)
sys.modules["gerar_estudo_cobertura_anexo_corrigido"] = mod
spec.loader.exec_module(mod)

necessarias = [
    "executar_estudo",
    "executar_estudo_dois_sellouts",
    "executar_cobertura_dash",
    "gerar_comparacao_estudos",
]

faltando = [nome for nome in necessarias if not hasattr(mod, nome)]
if faltando:
    raise SystemExit("ERRO: funções ausentes no motor: " + ", ".join(faltando))

print("OK: motor carregado e funções principais encontradas.")
