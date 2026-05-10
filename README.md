# SPS.TaxDeterminationImporter

Este arquivo fornece orientações ao Claude Code (claude.ai/code) ao trabalhar com o código deste repositório.

## Compilação

Esta é uma solução .NET Framework 4.8 (somente Windows, x64). Compile com MSBuild ou Visual Studio:

```bash
msbuild SPS.TaxDeterminationImporter.sln /p:Configuration=Debug /p:Platform="Any CPU"
msbuild SPS.TaxDeterminationImporter.sln /p:Configuration=Release /p:Platform="Any CPU"
```

Não há testes automatizados no projeto. A validação é feita manualmente dentro de um ambiente SAP Business One em execução.

## Dependências externas (fora do NuGet)

O projeto referencia DLLs via HintPaths relativos que precisam existir na máquina do desenvolvedor:

| Referência | Caminho esperado a partir da raiz da solution |
|---|---|
| `Interop.SAPbobsCOM.dll` | `../../../../Alfa/DLL/` |
| `Interop.SAPbouiCOM.dll` | `../../../../Alfa/DLL/` |
| `SBO.Hub.dll` | `../../../../Util/DLL/` |

São assemblies de interop COM do SDK do SAP Business One mais um wrapper proprietário (`SBO.Hub`). A solution não compila sem eles.

## Arquitetura

Este é um **Add-on para SAP Business One** — uma aplicação Windows desktop que se conecta ao cliente SAP B1 via COM e estende o formulário de Determinação de Código de Imposto (formulário interno SAP `80401`).

**Projetos da solution:**

- `SPS.TaxDeterminationImporter` — Ponto de entrada (WinExe). Conecta ao SAP, chama `InitializeBLL.Initialize()`, inicia a thread de escuta de eventos e chama `Application.Run()`.
- `SPS.TaxDeterminationImporter.Core` — Biblioteca de classes com toda a lógica de negócio, acesso a dados, formulários e modelos.

**Estrutura de camadas dentro do Core:**

```
BLL/   — Lógica de negócio. TaxDeterminationBLL é a classe principal.
DAO/   — Acesso a dados. Scripts.cs seleciona o ResourceManager correto (SQL Server ou HANA).
Forms/ — Handlers de eventos da UI do SAP. Não são WinForms — herdam das classes base do SBO.Hub.
Model/ — Classes de modelo C# simples com [SBO.Hub.Attributes] para mapeamento ORM.
Enum/  — TaxKeyFieldTypeEnum mapeia os tipos inteiros de campos-chave do SAP para valores nomeados.
Views/ — Arquivos XML .srf que definem os layouts dos diálogos customizados (FrmImportLog, FrmRemoveTax).
```

## Padrões importantes

### Tratamento de eventos da UI do SAP

Os formulários em `Forms/` não são WinForms padrão. Eles herdam de `SystemForm` ou `BaseForm` do `SBO.Hub` e sobrescrevem `ItemEvent()`. O registro de eventos é configurado em `EventFilterBLL.SetDefaultEvents()` usando `EventFilterHelper`. Apenas as combinações formulário/evento registradas são disparadas. Os handlers do formulário `80402` estão comentados — reative-os em `EventFilterBLL` se necessário.

`f80401` injeta três botões customizados no formulário nativo `80401` do SAP no evento `et_FORM_LOAD`. Toda a lógica executa apenas quando `!ItemEventInfo.BeforeAction` (eventos pós-ação).

### Queries independentes de banco de dados

Todo o SQL é armazenado como recursos de string em dois arquivos `.resx`:
- `DAO/SQL.resx` — queries para SQL Server
- `DAO/Hana.resx` — queries para SAP HANA

`Scripts.SetResourceManager()` seleciona o correto na inicialização com base em `SBOApp.Company.DbServerType`. Todo acesso a queries passa por `Scripts.Resource.GetString("ChaveDaQuery")`.

Ao adicionar uma nova query: inclua a string em **ambos** `SQL.resx` e `Hana.resx` com a mesma chave.

### Ciclo de vida de objetos COM

Todos os objetos do SDK SAP (`TaxCodeDeterminationsTCDService`, `TaxCodeDeterminationTCD`, `ProgressBar`, etc.) devem ser explicitamente liberados com `Marshal.ReleaseComObject()` em blocos `finally`. Não fazer isso causa vazamentos de memória no processo do cliente SAP. Veja `TaxDeterminationBLL.ImportData()` para o padrão estabelecido.

### Formato do Excel para importação

Layout esperado das colunas (linha 1 = cabeçalho, linha 2+ = dados):

| A | B | C | D | E | F | G | H | I | G+3 … |
|---|---|---|---|---|---|---|---|---|---|
| Valor1 | Valor2 | Valor3 | Valor4 | Efetivo De | Efetivo Ate | *Nome da utilização* | IVA - *Utilização* | DA - *Utilização* | *(repete por utilização)* |

As colunas 7+ se repetem em grupos de 3 para cada utilização de imposto. O nome da utilização na célula de cabeçalho deve corresponder a um registro existente em `OUSG` (sem distinção de maiúsculas/minúsculas). A exportação (`ExportData`) gera arquivos neste mesmo formato.

### Cache de validação

`TaxDeterminationBLL` armazena listas de validação (parceiros de negócio, itens, códigos NCM, etc.) em campos estáticos. São preenchidos de forma lazy por execução de importação. `BusinesPartnerList` é sempre recarregado a cada importação; os demais são verificados por nulo e reutilizados entre chamadas na mesma sessão do processo.

### Tipos de campo-chave

`TaxKeyFieldTypeEnum` espelha a numeração interna do SAP para tipos de campos-chave. Os tipos `NcmCode`, `ItemGroup`, `CustomerGroup` e `SupplierGroup` requerem conversão de descrição para ID (`ConvertDescriptionToId`) antes de gravar no serviço SAP. Os demais tipos (BP, Item, Estado, Filial, UDF) são passados diretamente.
