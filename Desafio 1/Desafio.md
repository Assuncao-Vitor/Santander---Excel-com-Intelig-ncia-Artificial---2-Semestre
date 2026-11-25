# Ferramenta de Controle de Investimentos com Excel

Planilha desenvolvida no Excel para simular a acumulação de patrimônio, estimar dividendos mensais e sugerir a alocação de aportes entre diferentes segmentos de Fundos de Investimento Imobiliário (FIIs), conforme o perfil do investidor.

---

## 1. Perguntas que a ferramenta responde e onde aparece cada resposta

As respostas estão organizadas na aba `Planilha1`:

- **Qual a sugestão de valor máximo para investir (até 30% do salário)?**

  Aparece em `Planilha1!D17`, calculada a partir do salário em `D15` (`=D15*30%`). Com salário de R$ 4.200,00, a sugestão é de **R$ 1.260,00**.

- **Quanto de patrimônio será acumulado ao final do prazo escolhido?**

  Aparece em `Planilha1!D23`, considerando o aporte mensal (`D20` = R$ 200,00), o prazo (`D21` = 5 anos) e a taxa de rendimento mensal (`D22` = 1,08%). O resultado atual é de **R$ 16.760,80**.

- **Quanto esse patrimônio gera em dividendos mensais?**

  Aparece em `Planilha1!D24`, multiplicando o patrimônio acumulado pelo rendimento mensal da carteira (`rendimento_carteira` = 0,60%), resultando em **R$ 100,56/mês**.

- **Qual a evolução do patrimônio e dos dividendos em diferentes prazos?**

  Aparece na tabela Previsões (`Planilha1!B27:D32`):

  | Prazo | Patrimônio acumulado | Dividendos mensais |
  |---|---:|---:|
  | 2 anos | R$ 5.446,17 | R$ 32,68 |
  | 5 anos | R$ 16.760,80 | R$ 100,56 |
  | 10 anos | R$ 48.691,53 | R$ 292,15 |
  | 20 anos | R$ 225.409,80 | R$ 1.352,46 |
  | 30 anos | R$ 866.780,96 | R$ 5.200,69 |

- **Como distribuir o aporte mensal entre os tipos de FIIs conforme o perfil?**

  Aparece na tabela de alocação (`Planilha1!B39:D46`), que apresenta os percentuais sugeridos e os respectivos valores em reais para cada tipo de FII, considerando o perfil selecionado em `C35` (atualmente Agressivo).

---

## 2. Como as funções `VF` e `PROCV` são utilizadas

### Função `VF` (Valor Futuro)

Calcula o montante acumulado com juros compostos e aportes mensais constantes.

**Simulação principal (`Planilha1!D23`):**

```excel
=VF(taxa_mensal; qtd_anos*12; aporte*-1)
```

Converte o prazo em anos para meses, multiplicando por 12, e inverte o sinal do aporte para apresentar o patrimônio acumulado como um valor positivo.

**Tabela de previsões (`Planilha1!C28:C32`):**

```excel
=VF($D$22; $A28*12; $D$20*-1)
```

Aplica o mesmo cálculo para os diferentes períodos de investimento: 2, 5, 10, 20 e 30 anos.

### Função `PROCV` (Busca Vertical)

Busca automaticamente o percentual de investimento recomendado para cada tipo de FII, de acordo com o perfil selecionado.

**Tabela de alocação (`Planilha1!C40:C45`):**

```excel
=PROCV($C$35&"-"&B40; Planilha2!$A:$D; 4; FALSO)
```

Concatena o perfil selecionado em `$C$35` com o tipo de FII da linha, formando uma chave de pesquisa, como `Agressivo-PAPEL`.

Em seguida, procura essa chave na coluna A da `Planilha2` e retorna o percentual correspondente da quarta coluna (D), utilizando correspondência exata (`FALSO`).

Os valores monetários são calculados multiplicando o percentual pelo aporte mensal:

```excel
=C40*$C$36
```

---

## 3. Intervalos nomeados criados

Foram criados cinco intervalos nomeados, apontando para células da `Planilha1`:

| Nome do intervalo | Referência | Valor atual | Finalidade |
|---|---|---:|---|
| `aporte` | `Planilha1!$D$20` | R$ 200,00 | Valor investido mensalmente. |
| `patrimonio` | `Planilha1!$D$23` | R$ 16.760,80 | Patrimônio acumulado projetado. |
| `qtd_anos` | `Planilha1!$D$21` | 5 | Prazo da simulação em anos. |
| `rendimento_carteira` | `Planilha1!$D$16` | 0,60% | Taxa mensal de dividendos da carteira. |
| `taxa_mensal` | `Planilha1!$D$22` | 1,08% | Taxa mensal utilizada no cálculo do valor futuro. |

---

## 4. Percentuais de cada perfil e origem

Os percentuais configurados na tabela de referência da `Planilha2` (`A2:D20`) são:

| Tipo de FII | Conservador | Moderado | Agressivo |
|---|---:|---:|---:|
| PAPEL | 30% | 45% | 30% |
| TIJOLO | 55% | 35% | 35% |
| HÍBRIDOS | 5% | 5% | 10% |
| FOFs | 10% | 10% | 10% |
| DESENVOLVIMENTO | 0% | 3% | 10% |
| HOTELARIAS | 0% | 2% | 5% |
| **Total** | **100%** | **100%** | **100%** |

**Origem dos percentuais:** Os pesos foram configurados para representar diferentes distribuições entre os segmentos de FIIs, considerando características de risco e retorno.

- **Conservador:** concentra 55% em Tijolo e 30% em Papel, com 10% em FOFs e 5% em Híbridos.
- **Moderado:** prioriza Papel (45%) e Tijolo (35%), com pequenas alocações nos demais segmentos.
- **Agressivo:** mantém Tijolo (35%) e Papel (30%), ampliando a participação em Híbridos (10%), FOFs (10%), Desenvolvimento (10%) e Hotelarias (5%).

---

## 5. O que mudou em relação à ferramenta do Expert

Em relação à estrutura base apresentada na aula do Expert, a planilha apresenta as seguintes alterações:

- **Parâmetros de simulação:** salário de R$ 4.200,00, rendimento mensal da carteira de 0,60%, aporte mensal de R$ 200,00, prazo de 5 anos, taxa mensal de 1,08% e perfil Agressivo.
- **Matriz de perfis expandida:** inclusão de seis tipos de FIIs na `Planilha2`, com percentuais específicos para cada perfil.
- **Chave composta:** utilização da fórmula `=B3&"-"&C3` para combinar o perfil do investidor com o tipo de FII.
- **Conferência dos percentuais:** inclusão de fórmulas de soma em `F3`, `F9` e `F15` para verificar se cada perfil totaliza 100%.
- **Teste de busca:** inclusão de uma consulta na célula `I3:J3`, utilizando a chave `Moderado-TIJOLO` para testar o funcionamento do `PROCV`.

---

**Observação:** Os valores apresentados são projeções matemáticas baseadas nas taxas e nos aportes configurados. Não representam garantia de rentabilidade ou de recebimento de dividendos.
