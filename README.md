
#Código Auxiliar para hidrodinâmica III
Código gerado para facilitar o trabalho no preenchimento da planilha de hidro3

## Dependências:
- numpy: pip install numpy
- itertools: pip install more-itertools
- xlwings: pip install xlwings

### Atributos da classe
A classe contém alguns atributos principais:

- Inteiro estruturaEscolhida: Permite a escolha entre [1 (Lista de Adjacência)](#Lista-de-Adjacência), [2 (Vetor de Adjacência)](#Vetor-de-Adjacência) e [3 (Matriz de Adjacência)](#Matriz-de-Adjacência) 
- EstruturaDeDados *estruturaGrafo: Objeto da classe virtual [EstruturaDeDados](#Estrutura-de-Dados) que criará a estrutura 
- Booleano Peso: É verdadeiro se tiver peso
- Booleano Direcionado: É verdadeiro se for direcionado
