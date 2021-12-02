# -*- coding: utf-8 -*-
*** Settings ***
Documentation   Starter robot for the Beginners' course.
Library         RPA.Browser.Selenium    auto_close=${FALSE}
Library         RPA.HTTP
Library         RPA.Excel.Files
Library         RPA.PDF


*** Keywords ***
Open The Intranet Website
    Open Available Browser    https://robotsparebinindustries.com/      #Acessa o site da aplicação
    Maximize Browser Window     #Maximiza a janela do navegador

*** Keywords ***
Log IN
    Input Text    username  maria   #Insere o nome do usuário
    Input Password    password  thoushallnotpass    #Insere a senha do Usuário
    Submit Form     #Clica no botão para efetuar o login
    Wait Until Page Contains Element    id:sales-form       #Aguarda pelo próximo elemento

***** Keywords ***
Fill The Form Using The Data From The Excel File
    Open Workbook    SalesData.xlsx     #Abre o arquivo Excel
    ${sales_reps}=    Read Worksheet As Table    header=True    #Ler o arquivo Excel e atribuir os dados para a variável "sales_reps"
    Close Workbook    #Fecha o arquivo Excel
    
    #Abaixo, usaremos um FOR para ler cada linha da planilha e inserir as informações.
    #OBS: A chamada da Classe "Fill and Submit The Form Fro one Person" está sendo feito dentro dessa classe.
    FOR    ${sales_reps}    IN    @{sales_reps}
        Fill and Submit The Form For One Person    ${sales_reps}
    END 
    

***** Keywords ***
Fill and Submit The Form For One Person
    [Arguments]    ${sales_reps}
    #Instruções usadas para inserir dados#    
    Input Text    firstname    ${sales_reps}[First Name]
    Input Text    lastname    ${sales_reps}[Last Name]
    Input Text    salesresult    ${sales_reps}[Sales]
    Select From List By Value    salestarget    ${sales_reps}[Sales Target]
    Click Button    //button[@type="submit"]


***** Keywords ***
Download The Excel File
    Download    https://robotsparebinindustries.com/SalesData.xlsx    overwrite=True    
    #Realiza o download do arquivo excel

***** Keywords ***
Collect The Result
    Screenshot    //div[@class="alert alert-dark sales-summary"]    ${CURDIR}${/}sales_sumary.png
    #Tira um print do resultado da automação e cria o arquivo sales_sumary.png

***** Keywords ***
Export The Table as a PDF
    #Aguarda o elemento aparecer na página
    Wait Until Element Is Visible    id:sales-results       
    #Manda as informações da tabela para a variável
    ${sales_result_html}=    Get Element Attribute    id:sales-results     outerHTML        
    #Pega as informações da variável e cria um arquivo PDf contendo as informações
    Html To Pdf    ${sales_result_html}}    ${CURDIR}${/}sales_results.pdf

***** Keywords ***
Log Out and Close The Browser
    #Essa função clica para efetuar o Log out e Fecha o navegador
    Click Button    Log out
    Close Browser

*** Tasks ***
Insert the sales data for the week and export it as a PDF
    Open The Intranet Website       #Chama a atividade: Open The Intranet Website    
    Log In                          #Chama a atividade: Log In
    Download The Excel File         #Chama a atividade: Download The Excel File 
    Fill The Form Using The Data From The Excel File        #Chama a atividade: Fill The Form Using The Data From The Excel File
    Collect The Result              #Chama a atividade: Collect The Result
    Export The Table as a PDF       #Chama a atividade: Export The Tables as a PDF
    [Teardown]    Log Out and Close The Browser    #Chama a atividade: Log Out and Close The Browser
    #OBS: A instrução [Teardown] executa a função caso aconteça alguma falha durante o processo
