# Controle de Inspeções Periódicas

Este projeto foi desenvolvido em Visual Basic para apresentar um sistema de verificação de execução de inspeções rotineiras. Tendo como base que a planilha será utilizada por um usuário e um admininstrador. 

## 🚀 Começando

Essas instruções permitirão que você obtenha uma cópia do projeto em operação na sua máquina local para fins de desenvolvimento e teste.

### 📋 Pré-requisitos

```
=> MSOffice 2019 ou superior instalado;
=> Habilitar a execução de macros na planilha;
```

### 🔧 Pré-configurações

1 - Mantenha a pasta raiz do programa conforme renomeada abaixo afim de evitar erros de destinos.

![pasta raiz](https://github.com/Cerqlau/Controle_de_inspecoes_periodicas/assets/87389666/c13a8389-713d-4e29-aa92-93543be7ed22.jpg)

2 -  Senha para edição do projeto: "@Teste"

-> Efetue a configuração dos destinatários de mail e texto na sub enviar_email localizada no Module1

### ⚙️ Executando o programa

-> O código possui rotinas para efetuar a verificação das inspeções. Possui um deadtime de +1 dia após a data de vencimento, e 5 dias máximos para antecipação de inspeções. 

-> Possui atualização automática do calendário a partir da data inserida da célula "AB4", solicitando o botão de "Atualizar calendário.

-> Ao atualizar o calendário um backup do extrato do relatório será efetuado e salvo nas pastas adequadas

-> A planilha é interativa o usuário poderá efetuar dois clicks na célula de cada inspeção e desta forma o respectivo formulário de inspeção será iniciado. 

-> O formulário possui verificações para evitar inconcistências: Falta de elaborador, Falta de observação para os itens que não foram inspecionados, etc. 

->Após o salvamento do relatório será gerado um arquivo PDF com os dados descritos e salvos na pasta do mês correspondente.

-> Após o salvamento será enviado um email para os destinatários pré configurados nas rotinas.



### 📨 Distribuição

É possivel efetuar a distribuição para usuários que possuem o MS Office instalado em suas máquinas, desde que estes estejem autorizados a execução de macros no programa. 


## 📦 Desenvolvimento

Lauro Cerqueira
LinkdIn: https://www.linkedin.com/in/lauro-cerqueira-70473568/

Instagram : laurorcerqueira

## 🛠️ Construído com

* [Microssoft Office Excel](https://docs.microsoft.com/pt-br/office/client-developer/excel/excel-home)
* [Visual Basic for Applications](https://docs.microsoft.com/pt-br/office/vba/api/overview/)


## 📄 Licença

Este projeto está sob a licença MIT - veja o arquivo [LICENSE.md](https://github.com/usuario/projeto/licenca) para detalhes.

## 🎁 

* Conte a outras pessoas sobre este projeto 📢
* Convide alguém da equipe para uma cerveja 🍺 
* Obrigado publicamente 🤓.

