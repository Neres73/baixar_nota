# Download de Notas Fiscais em Lote — SAP + Portal Web

Script em Python que baixa em lote os PDFs das notas fiscais listadas numa planilha de ocorrências: as notas de serviço pelo SAP e as que não estão no SAP por um portal web. No fim, confere se cada linha da planilha tem o seu PDF na pasta.

Uso no meu trabalho, no back-office de Manutenção. O download das notas era feito à mão por 2 pessoas, levando cerca de 6 dias de trabalho por mês somando os lotes. Hoje uma pessoa só faz com o script, e o script leva metade do tempo do processo manual.

---

## Por que existe

Cada nota exigia o mesmo caminho: procurar a ocorrência no SAP, abrir a nota, salvar o PDF com o nome certo e, se a nota não estivesse no SAP, repetir a busca no portal web. Além do volume, havia armadilhas que faziam o trabalho manual dar erro:

- **Notas com o mesmo número** na planilha que são notas diferentes, e acabavam sobrescrevendo uma à outra
- **Arquivo salvo que não era PDF** (uma página de erro com extensão `.pdf`)
- **Busca do portal que demora a atualizar** e ainda mostra o resultado da busca anterior, o que leva a baixar a nota errada

---

## O que o script faz

Menu no terminal com seis opções:

```
1 - Baixar notas de SERVIÇO no SAP
2 - Baixar notas de MATERIAL no portal web
3 - Verificar planilha × pasta
4 - Fluxo completo (1 → 2 → 3)
5 - Rebaixar só o que está faltando
6 - Trocar a planilha de entrada
```

**Leitura da planilha**
- Encontra sozinho as colunas de **nota** e de **ocorrência** pelo nome do cabeçalho (em qualquer posição nas primeiras linhas) e mostra o que achou para você confirmar. Se não achar, você informa as letras das colunas
- Dá um nome único para cada linha: a 2ª nota com o mesmo número vira `12345 (2).pdf`, a 3ª vira `12345 (3).pdf`, e assim por diante
- Ignora linhas sem número (títulos, legendas) e mostra quais foram puladas

**1. SAP (notas de serviço)**
- Conecta no SAP já aberto via **SAP GUI Scripting** e consulta cada ocorrência na transação
- Captura o PDF que o SAP gera numa pasta temporária, **em segundo plano**, sem depender de clique na tela. Na primeira nota, descobre sozinho qual pasta o SAP usa e passa a vigiar só ela
- Ocorrência não encontrada no SAP é marcada como **material** e vai para a etapa 2
- Mostra o progresso a cada 10 notas, com média por nota e horário estimado de término
- Gera `_relatorio_nfse.xlsx`, colorido por status

**2. Portal web (notas de material)**
- Abre o portal com **Playwright** (Edge, perfil salvo) e busca cada ocorrência marcada como material
- **Antes de baixar, confere se a linha que aparece na tela é mesmo a ocorrência buscada**. Se a busca não atualizou a tempo, não baixa nada e registra o erro com print da tela
- Baixa o PDF pelo visualizador, mesmo quando ele abre dentro de um iframe
- Se o navegador cair, reabre e tenta a mesma nota de novo

**3. Verificação**
- Confere linha a linha se o PDF existe na pasta e se é um PDF válido (assinatura `%PDF-` no início do arquivo)
- Lista os PDFs da pasta que não têm linha na planilha
- Gera `_verificacao_nfse.xlsx` com a situação de cada linha

**5. Reprocessamento**
- Roda a verificação, apaga os PDFs corrompidos e baixa de novo só o que está faltando (SAP → portal), depois confere outra vez

---

## Instalação

Requisitos: **Windows**, Python 3.10+, SAP GUI com scripting habilitado e Microsoft Edge.

```bash
git clone https://github.com/Neres73/baixar_nota.git
cd baixar_nota

python -m venv .venv
.venv\Scripts\activate

pip install -r requirements.txt
playwright install msedge
```

## Configuração

Copie o modelo e preencha:

```bash
copy .env.example .env
```

| Variável | O que é |
|---|---|
| `ARQUIVO_LISTA` | Planilha de ocorrências do mês (dá para trocar pela opção 6 do menu) |
| `PASTA_DESTINO` | Pasta onde os PDFs e os relatórios são salvos (padrão: `./notas`) |
| `NIX_URL` | Página de busca de ocorrências do portal web |

Ajustes finos (tempos de espera, pastas vigiadas, nomes de coluna aceitos, modo de teste com poucas notas) ficam nas constantes do topo do `main.py`, todas comentadas.

## Uso

```bash
python main.py
```

Deixe o SAP aberto e logado antes de escolher a opção 1. Na primeira vez que a etapa do portal rodar, faça o login na janela do Edge que abrir; a sessão fica salva para as próximas. O script **não lê, guarda nem digita senha**.

---

## Saídas

| Arquivo | Conteúdo |
|---|---|
| `<nota>.pdf` / `<nota> (2).pdf` | Uma nota por linha da planilha |
| `_relatorio_nfse.xlsx` | Resultado de cada ocorrência no SAP e no portal (verde OK, amarelo erro, vermelho material) |
| `_verificacao_nfse.xlsx` | Situação de cada linha (OK, faltando, PDF corrompido) e PDFs fora da lista |
| `_diagnostico_nix/` | Prints da tela quando o portal falha |

---

## Limitações

- Só roda no Windows, porque depende do SAP GUI Scripting
- Os caminhos de tela do SAP e os seletores do portal são os do ambiente onde o script foi feito; outro ambiente exige ajustes
- A captura em segundo plano depende de o SAP gravar o PDF numa pasta temporária. Se não gravar, existe um método reserva por cliques na tela (desligado por padrão, porque ocupa o computador)

## Estrutura

```
├── main.py            # menu, leitura da planilha, etapas SAP / portal / verificação
├── requirements.txt
├── .env.example       # modelo de configuração
└── .gitignore         # ignora .env, PDFs, planilhas e relatórios
```

---

Desenvolvido por [Gustavo Neres](https://github.com/Neres73), com apoio do Claude (IA) para modelar a lógica, escrever e revisar o código. As regras e as validações vêm do processo real e foram verificadas em execução. Projeto compartilhado como portfólio: nenhuma URL, credencial, caminho ou dado da empresa está no repositório. Licença MIT.
