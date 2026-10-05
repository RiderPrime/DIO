# Simulador de Investimentos em FIIs

## Sobre o projeto

Este projeto consiste em uma ferramenta desenvolvida no Excel para simular investimentos mensais em Fundos de Investimento Imobiliário (FIIs).

A ferramenta permite informar o aporte mensal, o prazo do investimento, a taxa de rendimento mensal e o perfil do investidor. A partir dessas informações, calcula o patrimônio acumulado e uma estimativa de dividendos mensais.

O projeto foi desenvolvido como parte de um desafio prático de Excel, com foco em aplicar funções financeiras, referências, intervalos nomeados, validação de dados e busca de informações.

> **Aviso:** os valores utilizados na planilha são fictícios e têm finalidade exclusivamente educacional. Os percentuais apresentados não representam recomendação de investimento.

---

## Perguntas que a ferramenta responde

A ferramenta foi construída para responder cinco perguntas principais:

### 1. Quanto investir por mês?

O valor do aporte mensal é informado na aba **Simulador**, no campo **Aporte mensal**.

Também existe uma sugestão de aporte equivalente a **30% do salário informado**, utilizada apenas como referência para a simulação.

### 2. Por quantos anos investir?

O usuário informa o prazo em anos no campo **Prazo (anos)**.

A planilha converte esse período para meses para realizar os cálculos financeiros.

### 3. Qual a taxa de rendimento mensal?

O usuário informa a taxa no campo **Rendimento mensal**.

Esse valor é utilizado tanto no cálculo do patrimônio acumulado quanto na estimativa dos dividendos mensais.

### 4. Quanto de patrimônio será acumulado?

O patrimônio acumulado aparece no campo **Patrimônio acumulado**.

A planilha também apresenta uma projeção para diferentes períodos:

* 2 anos
* 5 anos
* 10 anos
* 20 anos
* 30 anos

### 5. Quanto será recebido de dividendos por mês?

A estimativa de dividendos mensais é calculada a partir do patrimônio acumulado e da taxa de rendimento mensal.

O resultado aparece no campo **Dividendos mensais**.

---

## Como os cálculos funcionam

### Função VF

A função `VF` (Valor Futuro) é utilizada para calcular o patrimônio acumulado a partir dos aportes mensais.

A lógica utilizada é:

```excel
=VF(taxa_mensal;prazo_anos*12;-aporte;0;0)
```

Os principais argumentos são:

* `taxa_mensal`: rendimento mensal utilizado na simulação;
* `prazo_anos*12`: transforma o prazo informado em anos em quantidade de meses;
* `-aporte`: valor investido mensalmente;
* `0`: patrimônio inicial;
* `0`: indica que o aporte ocorre no final do período.

O resultado representa o patrimônio acumulado ao final do período escolhido.

### Cálculo dos dividendos

Depois de calcular o patrimônio, a estimativa de dividendos mensais é obtida por:

```excel
=Patrimônio*taxa_mensal
```

Na planilha:

```excel
=C12*taxa_mensal
```

---

## Intervalos nomeados

Para deixar as fórmulas mais legíveis e facilitar a manutenção da planilha, foram utilizados intervalos nomeados.

| Nome          | Informação                       |
| ------------- | -------------------------------- |
| `salario`     | Salário mensal informado         |
| `aporte`      | Aporte mensal                    |
| `prazo_anos`  | Prazo do investimento em anos    |
| `taxa_mensal` | Taxa de rendimento mensal        |
| `perfil`      | Perfil de investidor selecionado |

Por exemplo, em vez de utilizar apenas referências como `C6` e `C8`, a fórmula pode ser escrita de maneira mais clara:

```excel
=VF(taxa_mensal;prazo_anos*12;-aporte;0;0)
```

Isso facilita a leitura e reduz a dependência da posição física das células.

---

## Distribuição do aporte por perfil

A ferramenta também permite selecionar um perfil de investidor por meio de uma lista suspensa.

Foram utilizados três perfis:

* **Conservador**
* **Moderado**
* **Arrojado**

Cada perfil possui uma distribuição diferente entre seis tipos de fundos:

* Tijolo
* Papel
* Híbrido
* FOF
* Desenvolvimento
* Agro

Ao alterar o perfil, os percentuais e os valores destinados a cada tipo de fundo são atualizados automaticamente.

---

## PROCV e chave composta

Para encontrar o percentual correto de cada tipo de fundo, foi utilizada uma **chave composta**.

A chave combina o perfil do investidor com o tipo de fundo.

Por exemplo:

```text
Moderado_Tijolo
Moderado_Papel
Moderado_FOF
```

Na aba **Apoio**, essa chave é armazenada junto com o percentual correspondente.

O `PROCV` utiliza essa chave para localizar o percentual:

```excel
=VLOOKUP($C$9&"_"&F10;Apoio!$A$2:$D$19;4;FALSE)
```

O processo funciona da seguinte forma:

```text
Perfil selecionado
        ↓
Perfil + tipo de fundo
        ↓
   Chave composta
        ↓
       PROCV
        ↓
Percentual encontrado
        ↓
Aporte × percentual
```

Dess

