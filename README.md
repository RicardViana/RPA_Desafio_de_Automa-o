# Microsoft Power Automate Desktop - RPA para Iniciantes

# Projeto final

Ao final do curso de **Microsoft Power Automate Desktop RPA para Iniciantes** (https://www.udemy.com/course/microsoft-power-automate-desktop-rpa-para-iniciantes/) foi proposto o projeto onde era requisitado o seguinte desenvolvimento:

1) Acessar o site http://books.toscrape.com/ e extrair o nome do livro e seu respectivo valor até a terceira página
2) Preencher o formulario https://forms.gle/Lodo8gBwQW4zGsLX7 com as informações extraida 
3) Exportar os dados extraidos para um pasta de trabalho do Excel 
4) Enviar um e-mail através do Outlook anexado a pasta de trabalho com os dados extraidos 

Onde todo esse processo deve ser feito através de automação desenvolvida com o auxilio de um software RPA

## Extraindo os dados 

A primeira etapa do desenvolvimento foi definir algumas variaveis:

![image](https://user-images.githubusercontent.com/62486279/153303397-4b592967-cde5-4772-8f3a-9c80665f7639.png)

Essas que vão auxiliar no desenvolvimento do projeto.

Com as variaveis  definidas devemos iniciar uma instância do navegador, onde para esse projeto foi utilizado o Microsoft Edge 

![image](https://user-images.githubusercontent.com/62486279/153303506-fc8c4335-3291-49ec-9085-d0f4bd735f6b.png)

![image](https://user-images.githubusercontent.com/62486279/153303681-9da6d6ab-c205-40fe-bd87-3b2648fbce87.png)

E a ação **extrair dados de página da Web**

![image](https://user-images.githubusercontent.com/62486279/153303831-6685f32a-6791-409e-94d7-7c97c65fe903.png)

Essa que habilita uma caixa de seleção

![image](https://user-images.githubusercontent.com/62486279/153305128-1e967c72-b2ee-4527-a9c4-12b36a79b185.png)

Facilitando o processo de selecionar os dados que precisamos  

## Preencher Formulario 

Para o preenchimento do formulario, a primeira etapa necessaria foi de criar uma nova guia com o link do formulario

![image](https://user-images.githubusercontent.com/62486279/153307189-adceeb7f-2bd4-4512-b255-98188211ec1a.png)

E após, construir um loop (For Each) para preenchimento dos campos, envio e inicio de um novo preenchimento 

![image](https://user-images.githubusercontent.com/62486279/153307289-ab0db5f8-e37a-4054-96c2-05120a364065.png)

E por fim, encerramos as instâncias que não são mais necessarias

![image](https://user-images.githubusercontent.com/62486279/153307617-0da1a189-3df9-49db-a81e-e533a7adce3d.png)

## Exportados dados para o Excel 

Para essa ção, temos a opção de Gravar na planilha do Excel

![image](https://user-images.githubusercontent.com/62486279/153307727-85fe209b-cb57-4175-b2a0-6029c458e5e0.png)

Onde é utilizado a variavel que contem os dados extraido para preencher as celulas do Excel

## Enviar e-mail com o arquivo em Excel 

Concluido a autamação, é usado a ação **enviar mensagem de email por meio do outlook**

![image](https://user-images.githubusercontent.com/62486279/153307822-d5641563-de70-4a99-9492-b45be516999f.png)

Onde alem do texto no corpo do e-mail, é anexado a planilha com os dados extraidos do site de livros 
