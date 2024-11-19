# 🔎 Busca de Artigos

Uma aplicação para realizar buscas de termos em títulos de artigos científicos utilizando a API CrossRef. Os resultados são apresentados de forma organizada em uma interface gráfica, permitindo salvá-los em formato Excel para análise posterior.

---

## 📜 Funcionalidades

- **Busca de artigos**: Localiza até 100 artigos científicos relacionados ao termo inserido.
- **Exibição de resultados**: Mostra título, autores, ano e link dos artigos em uma tabela.
- **Ordenação por ano**: Resultados organizados em ordem cronológica (mais recente para mais antigo).
- **Acesso aos artigos**: Permite abrir os links diretamente na aplicação.
- **Exportação para Excel**: Salva os resultados em um arquivo Excel formatado.

---

## 🖼️ Screenshots

### Tela Inicial
![Tela Inicial](/tela01.png)

### Tela de Resultados
![Tela Resultados](/tela02.png)

### Tela de Salvamento
![Tela Salvamento](/tela03.png)

### Resultado do Salvamento
![Resultado Salvamento](/tela04.png)

---

## 🛠️ Tecnologias Utilizadas

- **Python**:
  - `tkinter` para interface gráfica.
  - `requests` para integração com a API CrossRef.
  - `openpyxl` para exportação de dados para Excel.
  - `webbrowser` para abertura de links.

---

## 📝 Como Funciona

1. **Insira o termo desejado** no campo de busca.
2. Clique no botão **"Buscar"** para iniciar a pesquisa.
3. Visualize os resultados na tabela:
   - **Título do artigo**.
   - **Autor(es)**.
   - **Ano de publicação**.
   - **Link para o artigo completo**.
4. Para salvar os resultados:
   - Clique em **"Salvar"** e escolha o local do arquivo.
5. Clique duas vezes em um artigo para **abrir o link no navegador**.
