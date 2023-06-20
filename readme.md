# ![SheetFusionLogomLOW](https://github.com/pedrobits/SheetFusion/assets/70610289/fd66b5ef-1df5-4a2e-8953-40ee3b7045f4)

## Descrição do Projeto
Este projeto consiste em um script em Node.js para converter arquivos de planilha no formato .xlsx e .csv para um único arquivo .xlsx. O script verifica a pasta "planilhas" no diretório atual e combina os arquivos de planilha presentes nessa pasta em um único arquivo fazendo um "merge" entre as planilhas, criando uma nova planilha consolidada.

## Requisitos
- Node.js (v12 ou superior)
- Pacotes NPM:
  - fs
  - path
  - xlsx
  - csv-parser
  - glob
  - readline-sync

Certifique-se de ter o Node.js instalado em sua máquina antes de executar o script. Os pacotes necessários serão instalados durante o processo.

## Instalação
1. Clone este repositório em sua máquina local.
2. Navegue até o diretório do projeto.
3. Abra o terminal e execute o seguinte comando para instalar as dependências necessárias:
```shell
npm install
```

## Utilização
1. Certifique-se de ter os arquivos de planilha (.xlsx ou .csv) que deseja mesclar na pasta "planilhas" no diretório do projeto.
2. Abra o terminal e execute o seguinte comando:
```shell
npm start
```
3. Siga as instruções apresentadas pelo script:
    - Caso existam pelo menos dois arquivos de planilha válidos na pasta "planilhas", o script combinará os arquivos em uma única planilha. Será solicitado que você forneça um nome para a nova planilha.
    - Se a combinação for bem-sucedida, o arquivo resultante será salvo no diretório do projeto com o nome fornecido e a extensão .xlsx.
    - Caso não seja possível criar a planilha, verifique se os arquivos de entrada têm conteúdo válido.
    - Se a pasta "planilhas" não existir no diretório do projeto, ela será criada automaticamente.

## Observações
- Certifique-se de ter permissões de leitura e gravação no diretório do projeto e na pasta "planilhas" para executar o script corretamente.
- Os arquivos de planilha a serem combinados devem estar no formato .xlsx ou .csv.
- Certifique-se de fornecer um nome exclusivo para a nova planilha que não conflite com outros arquivos no diretório do projeto.

## Contribuição
Contribuições para aprimorar este projeto são sempre bem-vindas. Sinta-se à vontade para abrir um "issue" ou enviar um "pull request" com suas sugestões ou correções.

### Gostaria de fazer uma contribuição para comprarmos pão de queijo?
[![Donate with PayPal](https://raw.githubusercontent.com/stefan-niedermann/paypal-donate-button/master/paypal-donate-button.png)](https://www.paypal.com/donate/?hosted_button_id=RAEXD3CNP3PTA)

## Licença
Este projeto está licenciado sob a [MIT License](LICENSE).
