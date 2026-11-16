"""
Atualização automática pela pasta da rede.

Ao abrir, o programa olha a pasta de atualizações na rede (só leitura).
Se lá tiver uma versão diferente da instalada, copia os arquivos do programa
(.py e ícone) para o computador e reabre sozinho. Sem rede/VPN, ou se a pasta
não existir, abre normalmente com a versão que já está instalada.

O versao.txt da rede tem a versão na 1ª linha e, nas seguintes, a lista
"arquivo|tamanho" do que faz parte dela. Só atualiza quando todos os arquivos
da lista estão lá com o tamanho certo — se alguém estiver copiando a versão
nova para a rede naquele momento, espera a próxima vez.

A pasta padrão está em PASTA_REDE. Para usar outra, crie o arquivo
"origem_atualizacao.txt" em %LOCALAPPDATA%\\<Programa>\\ com o caminho dentro.

Só vale para o programa instalado (pasta com "python" ao lado de "app");
rodando pelo PyCharm não faz nada.
"""
import os
import shutil
import subprocess
import sys
import tempfile
import threading
from pathlib import Path

try:                                                  # o .env fica ao lado deste arquivo
    from dotenv import load_dotenv
    load_dotenv(Path(__file__).resolve().with_name(".env"))
except ImportError:
    pass
# pasta compartilhada de onde saem as atualizações (ex.: \\servidor\equipe\Atualizacoes)
PASTA_REDE = os.getenv("PASTA_ATUALIZACAO", "")
ARQUIVO_VERSAO = "versao.txt"
EXTENSOES = (".py", ".ico", ".txt")
NAO_COPIAR = {"requirements.txt", "origem_atualizacao.txt"}
ESPERA_REDE = 4          # segundos — fora da rede o Windows pode demorar a desistir


def _ler(caminho: Path) -> str:
    try:
        return caminho.read_text(encoding="utf-8").strip()
    except Exception:
        return ""


def _com_limite(funcao, segundos):
    resultado = {}
    t = threading.Thread(target=lambda: resultado.setdefault("v", funcao()), daemon=True)
    t.start()
    t.join(segundos)
    return resultado.get("v")


def _versao_e_lista(texto: str):
    linhas = [l.strip() for l in texto.splitlines() if l.strip()]
    if not linhas:
        return "", {}
    lista = {}
    for l in linhas[1:]:
        nome, _, tamanho = l.partition("|")
        if nome and tamanho.isdigit():
            lista[nome] = int(tamanho)
    return linhas[0], lista


def versao_instalada(script=None) -> str:
    pasta_app = Path(script or __file__).resolve().parent
    return _versao_e_lista(_ler(pasta_app / ARQUIVO_VERSAO))[0]


def origem(pasta_app: Path):
    base = pasta_app.parent
    personalizada = _ler(base / "origem_atualizacao.txt")
    if not (personalizada or PASTA_REDE):
        return None                                   # nenhuma pasta configurada: não atualiza
    return Path(personalizada or PASTA_REDE) / base.name


def verificar(script: str) -> None:
    pasta_app = Path(script).resolve().parent
    if not (pasta_app.parent / "python").is_dir():
        return                                        # não é o programa instalado (ex.: PyCharm)
    pasta_rede = origem(pasta_app)
    if pasta_rede is None:
        return
    texto = _com_limite(lambda: _ler(pasta_rede / ARQUIVO_VERSAO), ESPERA_REDE)
    nova, lista = _versao_e_lista(texto or "")
    # só atualiza para versão MAIS NOVA (formato AAAA.MM.DD-HHMM: dá para comparar como texto);
    # se a rede tiver uma mais antiga, não volta para trás
    if not nova or nova <= versao_instalada(script):
        return
    if _com_limite(lambda: _copiar(pasta_rede, pasta_app, nova, lista), 120):
        # reabre com os arquivos novos (o que já estava carregado na memória é da versão antiga)
        subprocess.Popen([sys.executable, "-E", str(Path(script).resolve())], cwd=str(pasta_app),
                         close_fds=True)
        os._exit(0)


def _copiar(pasta_rede: Path, pasta_app: Path, nova: str, lista: dict) -> bool:
    try:
        if lista:
            arquivos = [pasta_rede / nome for nome in lista]
            # cópia para a rede ainda pela metade? deixa para a próxima vez que abrir
            if any(not a.is_file() or a.stat().st_size != tamanho for a, tamanho in zip(arquivos, lista.values())):
                return False
        else:
            arquivos = [a for a in pasta_rede.iterdir()
                        if a.is_file() and a.suffix.lower() in EXTENSOES
                        and a.name.lower() not in NAO_COPIAR and a.name != ARQUIVO_VERSAO]
        arquivos = [a for a in arquivos if a.suffix.lower() in EXTENSOES and a.name.lower() not in NAO_COPIAR]
        if not arquivos:
            return False
        # 1º copia tudo para uma pasta temporária: se a rede cair no meio, nada é trocado
        with tempfile.TemporaryDirectory(dir=pasta_app.parent) as tmp:
            for a in arquivos:
                shutil.copy2(a, Path(tmp) / a.name)
            for a in arquivos:
                os.replace(Path(tmp) / a.name, pasta_app / a.name)
        shutil.rmtree(pasta_app / "__pycache__", ignore_errors=True)
        (pasta_app / ARQUIVO_VERSAO).write_text(nova, encoding="utf-8")   # por último
        return True
    except Exception:
        return False
