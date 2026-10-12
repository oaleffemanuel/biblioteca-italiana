# Biblioteca Italiana

**Leia italiano no seu nível.** Um catálogo gratuito com 222 leituras em italiano, do A1 aos clássicos, organizado em sete etapas (*tappe*). Inclui um plano de leitura em PDF e uma planilha de acompanhamento.

🌐 **Site:** https://oaleffemanuel.github.io/biblioteca-italiana/

Para quem só quer ler, basta o site. Não precisa de conta nem de saber usar o GitHub.

## O que tem aqui

| | |
|---|---|
| [Catálogo](https://oaleffemanuel.github.io/biblioteca-italiana/catalogo.html) | 222 obras com filtros por etapa, forma, fonte, áudio e duração. Marcação de "Quero ler" e "Lido" salva no navegador. |
| [Plano de leitura (PDF)](docs/downloads/plano-de-leitura-biblioteca-italiana.pdf) | Método (leitura extensiva, áudio, vocabulário, *riassunto in tre frasi*), as 7 etapas com obras reais, teste de nível e dicas para brasileiros. |
| [Diário de leitura (XLSX)](docs/downloads/diario-de-leitura-biblioteca-italiana.xlsx) | Planilha com o catálogo, painel de progresso, sessões, vocabulário com revisão espaçada e resumos. Funciona no Excel e no Google Planilhas. |
| [Dados abertos](docs/dados/) | O catálogo em CSV e JSON. |

### As sete etapas

| Tappa | Nome | Nível | Obras |
|---|---|---|---:|
| 1 | *Primi passi* | A1 | 7 |
| 2 | *Prendere il ritmo* | A1–A2 | 22 |
| 3 | *Il ponte* | A2–B1 | 9 |
| 4 | *Il sentiero* | B1–B2 | 15 |
| 5 | *Primi originali* | B1+ (estimativa) | 11 |
| 6 | *La grande letteratura* | B2+ (estimativa) | 148 |
| 7 | *I classici* | C1+ (estimativa) | 10 |

## Direitos autorais

Acesso gratuito não é permissão para redistribuir. Por isso:

- A maior parte do catálogo é formada por **links para as páginas oficiais** de cada fonte: Quick Italian, Learn Italian Online, Liber Liber, Project Gutenberg, Wikisource e CLiGS.
- **Arquivos hospedados aqui:** apenas os 90 contos do corpus [CLiGS textbox](https://github.com/cligs/textbox), licenciados em CC BY-SA 4.0 (licença declarada em cada TEI original). Os PDFs em [`docs/textos/`](docs/textos/) foram gerados a partir desse TEI, sem alterar o texto, e seguem a mesma licença.
- **Áudios:** não hospedamos nenhum. Eles ficam nas páginas das fontes.
- **Obras com dúvida de direitos** aparecem só como link e marcadas como "em verificação". Uma adaptação de obra ainda protegida foi excluída.

É titular de direitos e quer corrigir ou retirar algo? [Abra uma issue](https://github.com/oaleffemanuel/biblioteca-italiana/issues).

## Licenças deste projeto

- **Código** (HTML, CSS, JS, scripts): [MIT](LICENSE).
- **Textos do site, plano de leitura, planilha e dados do catálogo:** [CC BY 4.0](https://creativecommons.org/licenses/by/4.0/deed.pt_BR).
- **PDFs em `docs/textos/`:** [CC BY-SA 4.0](https://creativecommons.org/licenses/by-sa/4.0/deed.pt_BR). Crédito: CLiGS textbox — *Collection of Italian Short Stories and Novellas (1880s-1920s)*. Veja [`docs/textos/CREDITOS.md`](docs/textos/CREDITOS.md).
- **Obras listadas com link:** seguem a licença de cada fonte.

## Para quem mantém o projeto

O site é HTML, CSS e JavaScript puro, sem framework e sem build no servidor, publicado pelo GitHub Pages a partir da pasta `docs/`.

```
data/catalogo.json         fonte única de verdade do catálogo
ferramentas/construir.py   gera as páginas, o PDF do plano, a planilha, o CSV/JSON e o sitemap
docs/                      o site publicado
  assets/                  estilo, script, ícone, imagem de compartilhamento
  downloads/               plano (PDF) e planilha (XLSX)
  textos/                  os 90 PDFs CLiGS (CC BY-SA 4.0)
  dados/                   catálogo em CSV e JSON
```

Para alterar o catálogo ou os textos:

```bash
pip install -r ferramentas/requirements.txt   # weasyprint e openpyxl
python3 ferramentas/construir.py              # use --sem-pdf para pular o PDF
```

Edite `data/catalogo.json` (obras) ou `ferramentas/construir.py` (textos das páginas e do plano), rode o script e faça o commit da pasta `docs/`.

---

Projeto pessoal e sem fins lucrativos. Não é afiliado a nenhuma das fontes citadas. *Buona lettura!*
