# 📊 Compilador de NFS-e Nacional (`XML` $\rightarrow$ `Excel Dinâmico`)

![Python](https://img.shields.io/badge/Python-3.10%2B-3776AB?style=for-the-badge&logo=python&logoColor=white)
![OpenPyXL](https://img.shields.io/badge/Excel-OpenPyXL-217346?style=for-the-badge&logo=microsoft-excel&logoColor=white)
![SPED NFS-e](https://img.shields.io/badge/Padr%C3%A3o_SPED-NFS--e_v1.01-1F4E79?style=for-the-badge)
![License](https://img.shields.io/badge/Licen%C3%A7a-MIT-green?style=for-the-badge)

Aplicação desktop e CLI em Python para leitura, auditoria fiscal e consolidação de lotes de **Notas Fiscais de Serviços Eletrônicas (NFS-e)** emitidas no **Padrão Nacional SPED (`nfse.gov.br`)**, gerando pastas de trabalho **Excel (`.xlsx`)** estruturadas com **somas e contagens 100% dinâmicas sob filtro (`SUBTOTAL`)**, agrupamento funcional de colunas e demonstrativos executivos.

---

## 🎯 Principais Diferenciais

- **Interface Gráfica (`Tkinter`) com Pré-Leitura Inteligente (Auto-Scan):**
  - Ao selecionar uma pasta com arquivos `.xml`, o sistema inspeciona automaticamente o lote em milissegundos, exibindo em cards de diagnóstico:
    - Quantidade de notas fiscais válidas;
    - Empresa principal identificada (Tomador ou Prestador predominante);
    - Número de fornecedores distintos e período de competência;
    - Valor bruto total dos serviços.
  - **Sugestão Automática de Nome:** Preenche automaticamente o nome ideal para a planilha `.xlsx` com base na empresa ou no diretório analisado.
- **Somas e Contagens Dinâmicas Reativas a Filtros (`SUBTOTAL`):**
  - Diferente de planilhas estáticas que usam `=SUM(...)`, todas as **15 colunas financeiras e tributárias** utilizam `=SUBTOTAL(109, ...)` e `=SUBTOTAL(103, ...)`.
  - Ao filtrar qualquer critério no Excel (ex.: *Tipo Retenção ISS = `2 - Retido na Fonte`*, *Retenção INSS*, *Fornecedor*, *Município*, *Competência* ou *Obra*), **todas as somas de ISSQN, INSS, IRRF, CSLL, PIS, COFINS, IBS, CBS, Valor Bruto e Valor Líquido recalculam instantaneamente** considerando apenas as notas filtradas.
- **Reorganização Funcional de 70 Colunas + Grupos Retráteis (`+` / `-`):**
  - As colunas críticas para conferência fiscal e financeira ficam visíveis logo na abertura, enquanto dados cadastrais extensos (endereços, CEPs, telefones, e-mails e chaves de acesso) ficam agrupados com recurso nativo de *Outline Grouping* do Excel.
- **Blindagem de Codificação e Higienização (`UTF-8` / `CP1252`):**
  - Tratamento automático contra *mojibake* e caracteres de controle inválidos, garantindo acentuação impecável em todos os campos.

---

## 🏗️ Arquitetura e Fluxo de Processamento

```mermaid
flowchart LR
    A["Lote de XMLs\n(nfse.gov.br v1.01)"] --> B["Auto-Scan & Parser XML\n(UTF-8 / CP1252)"]
    B --> C["Motor Fiscal & Contábil\n(ISSQN, Federais, IBS/CBS, Duplicatas)"]
    C --> D["Gerador Excel OpenPyXL\n(4 Abas + SUBTOTAL Dinâmico)"]
    D --> E["Planilha Executiva (.xlsx)\n& Exportação (.csv opcional)"]
```

---

## 📑 Estrutura da Planilha Gerada (4 Abas)

| Aba | Conteúdo e Funcionalidades |
| :--- | :--- |
| **1. `Resumo Financeiro Consolidado`** | Dashboard executivo com KPIs do lote, Demonstrativo Financeiro vinculado por fórmulas dinâmicas (`SUM`, `SUMIF`, `COUNTIF`), apuração da Reforma Tributária (**IBS / CBS**) e ranking **Top 10 Maiores Fornecedores**. |
| **2. `Notas Fiscais`** | Base analítica com **70 colunas** divididas em **7 blocos funcionais coloridos**, painel congelado (`D2`), destaque visual condicional para retenções na fonte e linha de **Total Filtrado** com `=SUBTOTAL(109, ...)`. |
| **3. `Resumo por Fornecedor`** | Consolidação por prestador de serviço com quebra explícita entre **ISSQN Retido na Fonte** vs. **ISSQN Não Retido**, retenções federais (`INSS`, `IRRF`, `CSLL`, `PIS`, `COFINS`), total retido e líquido a pagar. |
| **4. `Resumo por Município`** | Consolidação geográfica por município de execução do serviço, separando ISSQN retido, não retido e retenções federais com subtotais dinâmicos. |

### Organização dos 7 Blocos da Aba `Notas Fiscais`
1. **Identificação Essencial (`#1F4E79`):** `Número NFS-e`, `Competência`, `Razão Social Prestador`, `CNPJ Prestador`, `Data Emissão DPS`, `Status Duplicidade`, `Status (cStat)`.
2. **Enquadramento e Local (`#2F5597`):** `Regime Tributário (Simples Nacional)`, `Regime Especial`, `Município Prestação`, `Município Incidência`, `Cód. Obra`.
3. **Valores e ISS / ISSQN (`#1B6B53`):** `Tipo Retenção ISS`, `Valor Serviços (R$)`, `Deduções/Reduções (R$)`, `Desconto Incond. (R$)`, `Base Cálculo ISS (R$)`, `Alíquota ISS (%)`, `Valor ISSQN (R$)`.
4. **Retenções Federais e Valor Líquido (`#833C0C`):** `Retenções Federais Aplicadas` (filtro rápido), `Retenção INSS`, `Retenção IRRF`, `Retenção CSLL`, `Valor PIS`, `Valor COFINS`, `Total Retenções (R$)`, `Valor Líquido (R$)` *(+ detalhes de CST, Base e Alíquotas recolhíveis)*.
5. **Reforma Tributária — Novo SPED (`#4A3B7C`):** `IBS Alíquota UF (%)`, `IBS Alíquota Mun (%)`, `CBS Alíquota (%)`, `IBS Total (R$)`, `CBS Total (R$)`.
6. **Detalhamento do Serviço e Obra (`#2B5868`):** `Discriminação dos Serviços`, códigos e descrições de tributação nacional/municipal/NBS, informações complementares e endereço da obra.
7. **Cadastros Auxiliares e Chaves (`#4A5568` — Recolhível `+`/`-`):** Dados cadastrais completos do prestador e do tomador, códigos IBGE, números DFSe/DPS, chave de 50 dígitos e nome do arquivo XML.

---

## 🚀 Como Utilizar

### Pré-requisitos (Execução via Código-Fonte)
```bash
pip install -r requirements.txt
```

### 1. Modo Interface Gráfica (GUI)
Execute diretamente o script Python (ou dê duplo clique em `Iniciar_Compilador_NFSe.bat`):
```bash
python compilar_nfse.py
```

### 2. Modo Linha de Comando (CLI / Automação)
Você também pode compilar lotes via terminal passando a pasta de entrada (teste rápido com o XML fictício da pasta `examples/`):
```bash
python compilar_nfse.py --cli --input "examples" --output "demonstracao_nfse.xlsx"
```

---

## 📦 Gerando o Executável Único (`.exe`) para Distribuição

Para distribuir o aplicativo para outros computadores Windows sem necessidade de instalar Python:

```bash
pip install pyinstaller
python -m PyInstaller --noconfirm --onefile --windowed --name "Compilador_NFSe" compilar_nfse.py
```
O executável independente **`Compilador_NFSe.exe`** será gerado na pasta `dist/`.

---

## 🔒 Privacidade e Sigilo Fiscal (`.gitignore`)

Este repositório está configurado por padrão para **ignorar qualquer arquivo `.xml` real, pastas de lotes fiscais e planilhas `.xlsx`/`.csv` geradas**, garantindo que nenhum dado sensível de contribuintes seja versionado acidentalmente. Apenas o arquivo fictício em [`examples/exemplo_nfse_nacional.xml`](examples/exemplo_nfse_nacional.xml) é incluído para fins de demonstração.

---

## 👤 Autor

Desenvolvido por **Icaro** ([`@icaro-arch`](https://github.com/icaro-arch)).
