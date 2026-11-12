# App de Parte Diária de Equipamentos

App para celular em que o apontador da obra lança a parte diária dos equipamentos pesados. Ele substitui a ficha de papel, e todo dia os relatórios chegam por e-mail, prontos para assinatura e para a apropriação.

Feito em **Google Apps Script**, com uma **Planilha Google** como base de dados e o **Gmail** para os envios automáticos. Não precisa de servidor nem de mensalidade.

## Teste o app

**[Abrir a demonstração](https://leticiacristinaam-max.github.io/app-parte-diaria/)** (funciona no celular e no computador)

Para entrar, use o PIN **0000** (administradora) ou **1234** (apontador). A demonstração usa equipamentos e partes fictícios, guardados só no seu navegador, e não envia e-mails. O botão *Recomeçar*, no topo, volta aos dados iniciais.

## O que o app faz

- **Apontamento pelo celular:** o apontador escolhe o equipamento e informa o horímetro ou o KM, os horários de início e término, o abastecimento, as paradas com o motivo e a rua ou serviço do dia.
- **Horímetro inicial automático:** já vem preenchido com o final do dia anterior.
- **Pendências do dia:** mostra quais equipamentos ainda faltam lançar.
- **Funciona sem sinal:** se a internet cair, a parte fica guardada no celular e é enviada quando o sinal volta.
- **Cálculos automáticos:** jornada, intervalo, horas paradas (HP), horas disponíveis (HD) e horas efetivas (HT).
- **E-mail diário às 08h00, de segunda a sexta:**
  - para o gerente, o PDF com uma folha por equipamento, com campos de assinatura do apontador, do operador, do encarregado e do gerente, e o resumo do dia no final;
  - para a engenharia, o mesmo PDF e o **BDE em Excel** do mês, com HT/HD/HP dia a dia para colar na planilha de apropriação.
  - Na segunda-feira chega a parte de sexta. Sábado e domingo não há envio.
- **Relatório por período:** filtra por empresa ou equipamento e envia o PDF consolidado e o Excel para o e-mail da engenharia.
- **Cadastro de equipamentos:** prefixo, descrição, medidor (h ou km), operador, empresa, valor mensal e datas de entrada e saída da obra.
- **Dois níveis de acesso por PIN:** o **apontador** lança as partes. A **administradora** também cadastra equipamentos, altera os parâmetros, vê os valores mensais e envia relatórios.

## Arquivos

| Arquivo | O que é |
|---|---|
| `src/Code.gs` | Servidor: leitura e gravação na planilha, cálculos, geração do PDF e do Excel e envio dos e-mails |
| `src/Index.html` | Tela do app, com HTML, CSS e JavaScript |
| `docs/index.html` | Versão de demonstração que roda sozinha no navegador (GitHub Pages) |

## Como instalar

1. Crie uma **Planilha Google** nova e abra **Extensões → Apps Script**.
2. Cole o conteúdo de `src/Code.gs` no arquivo `Code.gs`.
3. Crie um arquivo HTML chamado **`Index`** e cole o conteúdo de `src/Index.html`.
4. Selecione a função **`configurar`**, clique em **Executar** e autorize. Ela cria as abas `Equipamentos`, `Partes` e `Config` e agenda o e-mail diário.
5. Clique em **Implantar → Nova implantação → App da Web**, com *Executar como: eu* e *Quem pode acessar: qualquer pessoa*.
6. Abra o link gerado no celular do apontador.
7. Na aba **Config** da planilha (ou na tela *Equipamentos* do app, como administradora), troque os **PINs** e os **e-mails**. Os valores do código são só exemplos.

Para testar o envio sem esperar o dia seguinte, rode a função `testarEnvio`. Ela manda o e-mail só para a engenharia.

## Observação

Esta versão do código não traz dados reais de obra. Os e-mails, PINs e equipamentos do código são exemplos.
