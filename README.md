Esse desafio tem como objetivo criar um script para automatizar um processo manual

Abaixo esta o passo-a-passo para criação e publicação do desafio

1. O script pode ser feito com a linguagem de sua preferencia, inclusive VBA

2. O script deve perguntar ao usuario em qual pasta deve buscar os arquivos

3. O script deve percorrer a pasta e abrir os arquivos .xls ou .xlsx que iniciam com o nome **RE_Química**

4. Os dados dos arquivos em excel devem ser convertidos para o formato .txt, respeitando o padrao do arquivo de exemplo _RE_Química_0001_06_01_2025_

   1. ![exemplo](https://github.com/user-attachments/assets/536c5929-7610-43e4-9a7c-9e9e8703c7e0)

   2. Todo arquivo se inicia com o cabeçalho `p  Nome_do_cliente`

   3. Toda fazenda se inicia com a linha `f	Nome_da_Fazenda`

   4. Toda amostra se inicia com a linha `a	Numero_da_Amostra_em_inteiro	Ano	Data_no_formato_DDMMAAAA`

   5. A sequencia de resultados é:
      | Linha | TXT | Excel |
      |:------------- | ------------- |:-------------:|
      r | 01MO | M.O.
      r | 01PHCACl2 | pH CaCl²
      r | 01PRES | P
      r | 01K | K
      r | 01CA | Ca
      r | 01MG | Mg
      r | 01AL | Al
      r | 01H+Al | H+Al SMP
      r | 01S | S
      r | 01H | (H+Al - Al)
      r | 01SB | S.B
      r | 01CTC | CTC
      r | 01V | Sat. Bases V%
      r | 01M | Sat. Al m%
      r | 01KCTC | (K / CTC)
      r | 01CACTC | (Ca / CTC)
      r | 01MGCTC | (Mg / CTC)
      r | 01CAMG | (Ca / Mg)

5. As linhas **01H, 01KCTC, 01CACTC, 01MGCTC, 01CAMG** são operações matematicas feitas com outros resultados

6. Todos os resultados devem estar no txt

7. O arquivo .txt deve ser salvo com o mesmo nome do arquivo referente de excel com a data no final

8. No repositorio, deve conter um executavel do script e o codigo implementado. Caso for feito em VBA, apenas o arquivo em excel com código dentro.

9. Ao finalizar, faça o commit do seu código e envie o link para felipe@athenasagricola.com.br
