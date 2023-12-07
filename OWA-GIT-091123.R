library(shiny)
library(shinyWidgets)
library(readxl)
library("PerformanceAnalytics")
library(fields)
library(RColorBrewer)
library(extrafont)
library(corrplot)
library(ggplot2)
library(reshape2)
library(openxlsx)
library("stats")
library("car")
library("ggplot2")
library("ggpubr")
library(shinyFiles)
library(plotly)

ui <- fluidPage(
  
  tags$head(
    tags$style(HTML(".info-icon { cursor: pointer; }")),
  ),
  
  # Define the title that will appear in the application
  # Define o título que aparecerá na aplicação
  titlePanel(
    tags$h2(
      HTML('<img src="https://cdn.cookielaw.org/logos/2182e5b0-78dc-4d97-a438-fb223b451736/842ea97a-a8f2-46e9-982c-eba6d612165c/677304d5-c81f-4506-976e-ab32a31ec753/logoPucMinas.png" width="100" height="100" style="float: right; margin-right: 10px;">OWA'),
      style = "color: #428bca; font-weight: bold;"
    )
  ),
  
  tabsetPanel(
    tabPanel("Data Input", 
             sidebarLayout(
               # Components for choosing the Excel file
               # Componentes para a escolha do arquivo Excel
               sidebarPanel(
                 div(
                   style = "display: flex; align-items: center;",
                   HTML("<span style='font-weight: bold;'>Select Excel file</span>"),
                   
                   div(
                     class = "info-icon-0",
                     actionLink(inputId = "info-dialog-0-link", label = icon("question-circle"), style = "background: none; border: none; padding: 0; margin-left: 5px;")
                   )
                 ),
                 fileInput(inputId = "file", label = NULL),
                 # Components for choosing sub-indicators
                 # Componentes para a escolha dos subindicadores
                 div(
                   style = "display: flex; align-items: center;",
                   HTML("<span style='font-weight: bold;'>Select columns</span>"),
                   
                   div(
                     class = "info-icon-1",
                     actionLink(inputId = "info-dialog-1-link", label = icon("question-circle"), style = "background: none; border: none; padding: 0; margin-left: 5px;")
                   )
                 ),
                 #br(),
                 pickerInput(
                   inputId = "selected_headers",
                   label = NULL,
                   choices = NULL,  
                   multiple = TRUE
                 ),
                 # Components for choosing the control variable
                 # Componentes para a escolha da variável de controle
                 div(
                   style = "display: flex; align-items: center;",
                   HTML("<span style='font-weight: bold;'>Select the control variable</span>"),
                   
                   div(
                     class = "info-icon-2",
                     id = "info-dialog-2",  # Novo ID para o segundo botão
                     actionLink(inputId = "info-dialog-2-link", label = icon("question-circle"), style = "background: none; border: none; padding: 0; margin-left: 5px;")
                   )
                 ),
                 selectInput(
                   inputId = "selected_header",
                   label = NULL,
                   choices = NULL,  
                   selected = NULL
                 ),
                 # Components for choosing emphasis (Positive or Negative)
                 # Componentes para a escolha da ênfase (Positiva ou Negativa)
                 div(
                   style = "display: flex; align-items: center;",
                   HTML("<span style='font-weight: bold;'>Choose the emphasis</span>"),
                   
                   div(
                     class = "info-icon-3",
                     actionLink(inputId = "info-dialog-3", label = icon("question-circle"), style = "background: none; border: none; padding: 0; margin-left: 5px;")
                   )
                 ),
                 radioButtons(
                   inputId = "positive_negative",
                   label = NULL,
                   choices = c("Positive", "Negative"),
                   selected = "Positive"
                 ),
                 # Components for choosing the emphasis percentage
                 # Componentes para a escolha da porcentagem da ênfase
                 div(
                   style = "display: flex; align-items: center;",
                   HTML("<span style='font-weight: bold;'>Select the emphasis percentage</span>"),
                   
                   div(
                     class = "info-icon-4",
                     actionLink(inputId = "info-dialog", label = icon("question-circle"), style = "background: none; border: none; padding: 0; margin-left: 5px;")
                   )
                 ),
                 br(),
                 sliderInput(
                   inputId = "custom_slider",
                   label = NULL,
                   min = 1,
                   max = 100,
                   value = 1
                 ),
                 # Button for calculating the composite indicator
                 # Botão para o cálculo do indicador composto
                 actionButton(inputId = "calcular_owa", label = "Calculate"),
                 div(
                   style = "display: flex; align-items: center;",
                   HTML("<span style='font-weight: bold;'></span>"),
                 ),
                 br(),
                 # Button to download xlsx that has step-by-step calculations to find the composite indicator
                 # Botão para o download do xlsx que possui o passo a passo dos cálculos para encontrar o indicador composto
                 div(
                   style = "display: flex; align-items: center;",
                   HTML("<br><span style='font-weight: bold;'> </span>"),
                   actionButton(inputId = "download_button", label = "Download xlsx"),
                   
                   div(
                     class = "info-icon-6",
                     actionLink(inputId = "info-dialog-6-link", label = icon("question-circle"), style = "background: none; border: none; padding: 0; margin-left: 5px;")
                   )
                 ),
               ),
               
               mainPanel(
                 verbatimTextOutput("selected_output"),
               )
             )),
    # Components to show correlation chart 1
    # Componentes para mostrar o gráfico 1 de correlação 
    tabPanel("Graph 1 - Correlation Matrix",
             mainPanel(
               plotOutput("correlationPlot", width = "800px", height = "600px"),
               div(
                 downloadButton("download_plot1", "Download Graph 1"),
                 style = "margin-bottom: 10px;"
               )
             )
    ),
    # Components to show outlier graph 2
    # Componentes para mostrar o gráfico 2 de outliers
    tabPanel("Graph 2 - Ratio of atypical measurements",
             mainPanel( width = 8,
                        plotOutput("scatterplot", width = "800px", height = "600px"),
                        div(
                          downloadButton("download_plot2", "Download Graph 2"),
                          style = "margin-bottom: 10px;"
                        )
             )
    ),
    # Components to show explanatory power chart 3
    # Componentes para mostrar o gráfico 3 de poder explicativo 
    tabPanel("Graph 3 - Explanatory Power",
             mainPanel( width = 8,
                        plotOutput("correlation_plot", width = "800px", height = "600px"),
                        div(
                          downloadButton("download_plot3", "Download Graph 3"),
                          style = "margin-bottom: 10px;"
                        )
             )
    ),
    # Components to show uncertainty graph 4
    # Componentes para mostrar o gráfico 4 de incerteza
    tabPanel("Graph 4 - Uncertainty (Ranking Variation)",
             mainPanel( width = 8,
                        plotlyOutput("graficoCandlestick"),
                        div(
                          downloadButton("download_plot4", "Download Graph 4"),
                          style = "margin-bottom: 10px;"
                        )
             )
    ),
    # Components to show discriminant power graph 5
    # Componentes para mostrar o gráfico 5 de poder discriminante 
    tabPanel("Graph 5 - Discriminating Power",
             mainPanel( width = 8,
                        plotOutput("histograma", width = "800px", height = "600px"),
                        div(
                          downloadButton("download_plot5", "Download Graph 5"),
                          style = "margin-bottom: 10px;"
                        )
             )
    ),
    
    # Components to help the user when clicking on question mark icons
    # Componentes para ajudar o usuário ao clicar nos ícones de interrogação
    tags$script(
      HTML('
      $(document).ready(function(){
        $(".info-icon-1").click(function(){
          alert("Select the columns you want to analyze.");
        });

        $(".info-icon-2").click(function(){
          alert("The control variable must be the column that has the greatest association with the problem.");
        });
        
        $(".info-icon-3").click(function(){
          alert("The chosen emphasis means focusing on larger (positive) or smaller (negative) values.");
        });
        
        $(".info-icon-4").click(function(){
          alert("The percentage will define the number of values that will be analyzed according to the chosen emphasis.");
        });
        
        $(".info-icon-0").click(function(){
          alert("The file must start from line 2 and already have its data normalized.");
        });
        
        $(".info-icon-6").click(function(){
          alert("For a better visualization and understanding of the calculations and processes carried out, an xslx was created. To save it on your computer just click the Download .xlsx button");
        });
        
      });
    ')
    ),
  )
)

server <- function(input, output, session) {
  
  ordenar_eq <- function(x) {
    order(order(x))
  }
  
  headers <- reactive({
    req(input$file)  
    excel_data <- read_excel(input$file$datapath, sheet = 1)
    colnames(excel_data)
  })
  
  observe({
    updatePickerInput(
      session,
      inputId = "selected_headers",
      choices = headers()
    )
    updateSelectInput(
      session,
      inputId = "selected_header",
      choices = headers()
    )
  })
  
  calcular_owa_function <- function() {
    
    selected_headers <- input$selected_headers
    selected_header <- input$selected_header
    selected_pn <- input$positive_negative
    selected_slider <- input$custom_slider
    
    # Name of the file available for downloading the step-by-step calculations of the composite indicator
    # Nome para o arquivo disponível para o download do passo a passo dos cálculos do indicador composto
    data_de_hoje <- Sys.Date()
    data_formatada <- format(data_de_hoje, "%d-%m-%Y")
    nome_do_arquivo <- paste0("Passo a passo ",data_formatada,".xlsx")
    
    # Check if the file already exists and delete it
    # Verificar se o arquivo já existe e excluí-lo
    if (file.exists(nome_do_arquivo)) {
      file.remove(nome_do_arquivo)
    }
    
    # Create a new Excel file
    # Criar um novo arquivo Excel
    wb <- createWorkbook()
    # Add tabs in Excel step by step
    # Adiciona abas ao Excel passo a passo
    addWorksheet(wb, "Dados")
    addWorksheet(wb, "Normalizada")
    addWorksheet(wb, "Dados escolhidos")
    addWorksheet(wb, "Transposta")
    addWorksheet(wb, "Ordenada")
    addWorksheet(wb, "OWA")
    addWorksheet(wb, "Indicador Composto")
    
    # Read the spreadsheet
    # Ler a planilha
    dados <- read_excel(input$file$datapath, sheet = 1)
    head(dados)
    
    # Define the control variable
    # Define a variável de controle
    variavel_controle <- dados[c(selected_header)]
    lista_variavel_controle <- as.list(variavel_controle)
    
    # Convert spreadsheet data to a matrix
    # Converter os dados da planilha para uma matriz
    matriz_dados <- as.matrix(dados)
    matriz_dados_normalizada <- matriz_dados
    
    # Data normalization
    # Normalização de dados
    num_cols <- ncol(matriz_dados)
    
    for (col in 1:num_cols) {
      if (any(grepl("^[a-zA-Z]+$", matriz_dados[, col]))) {
        matriz_dados_normalizada[,col] <- matriz_dados[,col]
      } else {
        
        # Defines the position of the column that corresponds to the control variable
        # Define a posição da coluna que corresponde a variável de controle
        posicao_coluna_variavel_controle <- which(colnames(matriz_dados) == selected_header)
        coluna_variavel_controle <- matriz_dados[, posicao_coluna_variavel_controle]
        # Converts the values of the control variable into a list
        # Converte em uma lista os valores da variável de controle
        lista_valores_variavel_constrole <- as.numeric(as.list(as.numeric(coluna_variavel_controle)))
        lista_valores_coluna <- as.numeric(as.list(as.numeric(matriz_dados[,col])))
        
        correlacao <- cor(lista_valores_variavel_constrole, lista_valores_coluna)

        if (correlacao < 0) {
          coluna_correlacao_positiva <- as.numeric(matriz_dados[,col]) * -1
        } else {
          coluna_correlacao_positiva <- as.numeric(matriz_dados[,col])
        }
        
        min_col <- as.numeric(min(coluna_correlacao_positiva))
        max_col <- as.numeric(max(coluna_correlacao_positiva))
        
        # Normalize only if column contains numbers
        # Normalizar apenas se a coluna contiver números
        matriz_dados_normalizada[,col] <- as.numeric((as.numeric(coluna_correlacao_positiva) - min_col) / (max_col - min_col))
        
      }
    }
    
    # Defines the matrix with the chosen subindicators
    # Define a matriz com os subindicadores escolhidos
    posicoes_colunas <- sapply(selected_headers, function(nome) which(colnames(matriz_dados_normalizada) == nome))
    my_data <- matriz_dados_normalizada[, posicoes_colunas]
    my_data_numeric <- apply(my_data, 2, function(x) as.numeric(x))
    planilha_grafico1 <- my_data_numeric
    
    # Write the matrix in the sheet tab
    # Escreve na aba sheet a matriz
    writeData(wb, sheet = "Dados", matriz_dados, startRow = 1, startCol = 1)
    writeData(wb, sheet = "Normalizada", matriz_dados_normalizada, startRow = 1, startCol = 1)
    writeData(wb, sheet = "Dados escolhidos", my_data, startRow = 1, startCol = 1)
    
    # Define the matrix with the chosen columns
    # Definir a matriz com as colunas escolhidas
    matriz_resultante <- my_data
    quantidade_de_elementos_coluna <- nrow(matriz_resultante)
    
    # Creating the matrix transpose
    # Criando a transposta da matriz
    matriz_transposta <- t(matriz_resultante)
    writeData(wb, sheet = "Transposta", matriz_transposta, startRow = 1, startCol = 1)
    
    # Defines the position of the column that corresponds to the control variable
    # Define a posição da coluna que corresponde a variável de controle
    posicao_coluna_variavel_controle <- which(colnames(matriz_dados_normalizada) == selected_header)
    coluna_resultante <- matriz_dados_normalizada[, posicao_coluna_variavel_controle]
    # Converts the values of the control variable into a list
    # Converte em uma lista os valores da variável de controle
    lista_ultima_linha <- as.list(coluna_resultante)
    lista_linha_variavel_controle <- lista_ultima_linha
    
    # Sort the transposed matrix that has the subindicators in descending order
    # Ordena a matriz transposta que possui os subindicadores de forma decrescente
    matriz_ordenada <- apply(matriz_transposta, 2, function(x) sort(x, decreasing = TRUE))
    # Write the ordered matrix in Excel
    # Escreve no excel a matriz ordenada
    quantidade_colunas_matriz_ordenada = ncol(matriz_ordenada)
    writeData(wb, sheet = "Ordenada", matriz_ordenada, startRow = 1, startCol = 1)
    
    # Sets weight based on emphasis choice
    # Define o peso baseado na escolha da ênfase
    quantidade_linhas <- nrow(matriz_ordenada)
    valor_inicial = 0
    somar <- 1/quantidade_linhas
    tamanho_lista = quantidade_linhas
    
    lista_valores <- list()
    
    for (i in 1:nrow(matriz_ordenada)) {
      lista_valores[[i]] <- valor_inicial * 100
      valor_inicial <- valor_inicial + somar
    }
    
    posicao_final <- nrow(matriz_ordenada) + 1
    lista_valores[[posicao_final]] <- 100
    
    selecionado <- 0
    for (i in 1:posicao_final){
      if (lista_valores[[i]] <= selected_slider) {
        selecionado <- i
        if (i+1 <= posicao_final) {
          selecionado <- i + 1
        }
      } else {
        break
      }
    }
    
    # List of possible emphases
    # Lista das ênfases possiveis
    if (selecionado == 0) {
      selecionado = lista_valores[1]
    }
    
    porcentagem = as.numeric(lista_valores[selecionado])
    
    # Starts at position 2 - I choose the minimum emphasis - I leave all the lines
    # Começa na posição 2 - escolho a ênfase mínima - deixo todas as linhas
    if (porcentagem == 0) {
      porcentagem = as.numeric(lista_valores[2])
    }
    
    # Set default values for comparison calculation (emphasis 0%)
    # Definimos valores padrão para o cálculo da comparação (ênfase 0%)
    porcentagem_fixa = as.numeric(lista_valores[2])
    peso_divisao_fixo = (porcentagem_fixa/100)/(1/quantidade_linhas)
    quantidade_linhas_excluidas_fixo = round((porcentagem_fixa/100) * quantidade_linhas) - 1
    sobram_linhas_fixo = quantidade_linhas - quantidade_linhas_excluidas_fixo

    matriz_peso_fixo <- matriz_ordenada
    
    linhas_ordenada_fixo = nrow(matriz_ordenada)
    colunas_ordenada_fixo = ncol(matriz_ordenada)
    
    linhas_peso_fixo = nrow(matriz_peso_fixo)
    colunas_peso_fixo = ncol(matriz_peso_fixo)
    
    quantidade_colunas_fixo = as.integer(ncol(matriz_ordenada))
    
    linhas_para_manter_fixo <- (nrow(matriz_peso_fixo) - quantidade_linhas_excluidas_fixo)
    nova_matriz_fixo <- matriz_peso_fixo[1:linhas_para_manter_fixo, ]
    
    if (porcentagem_fixa == 100 || linhas_para_manter_fixo == 1) {
      linhas_para_manter_fixo = 1
      nova_matriz_fixo <- t(matrix(nova_matriz_fixo))
    }
    
    peso_fixo = porcentagem_fixa/100
    sobram_linhas_fixo = (quantidade_linhas - linhas_para_manter_fixo) + 1
    
    peso_fixo = 1/linhas_para_manter_fixo
    
    n_coluna_fixo = quantidade_colunas_fixo
    n_linha_fixo = linhas_para_manter_fixo
    total_fixo <- 1:(n_linha_fixo * n_coluna_fixo)
    matriz_fixo <- matrix(total_fixo, nrow = linhas_para_manter_fixo, ncol = n_coluna_fixo)
    
    # Multiply each position in the matrix by the emphasis weight
    # Multiplicar cada posição da matriz pelo peso da ênfase
    for (i in 1:nrow(matriz_fixo)) {
      for (j in 1:ncol(matriz_fixo)) {
        matriz_fixo[i, j] <- as.numeric(nova_matriz_fixo[i, j]) * peso_fixo
      }
    }
    
    somas_colunas_fixo <- list()
    for (col in 1:ncol(matriz_fixo)) {
      soma_fixo <- sum(matriz_fixo[, col])
      somas_colunas_fixo[[col]] <- soma_fixo
    }
    
    # Composite indicator for comparison
    # Indicador composto para a comparação
    somas_colunas_fixo <- as.numeric(somas_colunas_fixo)
    ordem_eq_resultado_fixo <- ordenar_eq(somas_colunas_fixo)
    
    quantidade_linhas_excluidas = round((porcentagem/100) * quantidade_linhas) - 1
    sobram_linhas = quantidade_linhas - quantidade_linhas_excluidas
    
    # Delete rows in sorted array
    # Excluir linhas na matriz ordenada
    matriz_peso <- matriz_ordenada
    
    linhas_ordenada = nrow(matriz_ordenada)
    colunas_ordenada = ncol(matriz_ordenada)
    
    linhas_peso = nrow(matriz_peso)
    colunas_peso = ncol(matriz_peso)
    
    quantidade_colunas = as.integer(ncol(matriz_ordenada))
    
    # Checks the emphasis chosen for constructing the composite indicator
    # Verifica a ênfase escolhida para a construção do indicador composto
    if (selected_pn == "Positive") {
      nova_matriz <- matriz_peso[1:sobram_linhas, ]
      
      if (porcentagem == 100 || sobram_linhas == 1) {
        linhas_para_manter = 1
        sobram_linhas = 1
        nova_matriz <- matriz_peso[1, , drop = FALSE]
      }
      
    } else {
      nova_matriz <- tail(matriz_peso, n = sobram_linhas)
      
      if (porcentagem == 100 || sobram_linhas == 1) {
        linhas_para_manter = 1
        sobram_linhas = 1
        nova_matriz <- matriz_peso[nrow(matriz_peso), , drop = FALSE]
      }
      
      if (sobram_linhas == nrow(matriz_peso)) {
        nova_matriz <- matriz_peso
      }
      
    }
    
    # Calculate the weight according to the chosen emphasis percentage
    # Calcula o peso de acordo com a porcentagem da ênfase escolhida
    peso = porcentagem/100
    peso = 1/sobram_linhas
    n_coluna = quantidade_colunas
    n_linha = sobram_linhas
    total <- 1:(n_linha * n_coluna)
    matriz <- matrix(total, nrow = n_linha, ncol = n_coluna)
    
    # Multiply each position in the matrix by the emphasis weight
    # Multiplicar cada posição da matriz pelo peso da ênfase
    for (i in 1:nrow(matriz)) {
      for (j in 1:ncol(matriz)) {
        matriz[i, j] <- as.numeric(nova_matriz[i, j]) * peso
      }
    }
    
    writeData(wb, sheet = "OWA", matriz, startRow = 1, startCol = 1)
    
    num_linhas_owa_correlacao <- nrow(matriz)
    num_linhas_owa_correlacao <- num_linhas_owa_correlacao + 2
    num_linhas_owa_media <- num_linhas_owa_correlacao + 1
    
    # Find the composite indicator
    # Encontra o indicador composto
    somas_colunas <- list()
    for (col in 1:ncol(matriz)) {
      soma <- sum(matriz[, col])
      somas_colunas[[col]] <- soma
    }
    
    writeData(wb, sheet = "Indicador Composto", somas_colunas, startRow = 1, startCol = 1)
    
    # Calculate the average of the composite indicator
    # Calcula a média do indicador composto
    media_soma <- mean(as.numeric(somas_colunas))
    
    # Calculate the correlation between the control variable and the composite indicator
    # Calcular a correlação entre a variável de controle e o indicador composto
    lista_ultima_linha <- as.numeric(lista_ultima_linha)
    somas_colunas <- as.numeric(somas_colunas)
    ordem_eq_resultado <- ordenar_eq(somas_colunas)
    
    correlacao <- cor(lista_ultima_linha, somas_colunas)
    correlacao_resultado <- correlacao
    
    # Calculates the difference between the user-defined composite indicator and the default 0% emphasis indicator
    # Calcula a diferença entre o indicador composto definido pelo usuário e o indicador padrão de ênfase 0%
    diferenca = somas_colunas_fixo - somas_colunas
    diferenca = abs(ordem_eq_resultado_fixo - ordem_eq_resultado)
    media_diferenca <- mean(diferenca)
    
    # Calculate entropy
    # Calcula a entropia
    probabilities <- table(somas_colunas) / length(somas_colunas)
    lista <- as.list(probabilities * log2(probabilities))
    valores_vetor <<- as.numeric(lista)
    
    # Get the keys (names) from the list
    # Pega as chaves (nomes) da lista
    chaves <<- names(lista)
    # Get the values from the list
    # Pega os valores da lista
    valores <- unlist(valores_vetor)
    entropy <- -sum(probabilities * log2(probabilities))
    
    # Calculate the coefficient of variation
    # Calcula o coeficiente de variação
    media = mean(somas_colunas)
    desvio_padrao = sd(somas_colunas)
    cv = desvio_padrao/media
    
    # Etapa 1
    # Select only the variables
    # Transformar as listas em um data frame
    data <- data.frame(Coluna1 = somas_colunas, Coluna2 = lista_variavel_controle)
    # Definir nomes para as colunas da matriz
    colnames(data) <- c("OWA", "V_Controle")
    
    # Finding the center point
    data.center = colMeans(data)
    
    # Finding the covariance matrix
    data.cov = cov(data)
    
    # Etapa 2
    # Ellipse radius from Chi-Sqaure distrubiton
    rad  = qchisq(p = 0.95 , df = ncol(data))
    
    # Square root of Chi-Square value
    rad  = sqrt(rad)
    
    # Finding ellipse coordiantes
    ellipse <- car::ellipse(center = data.center , shape = data.cov , radius = rad ,
                            segments = 150 , draw = FALSE)
    
    
    # Etapa 3
    # Ellipse coordinates names should be same with air data set
    ellipse <- as.data.frame(ellipse)
    colnames(ellipse) <- colnames(data)
    
    # Create scatter Plot
    figure <- ggplot(data , aes(x = OWA, y = V_Controle)) +
      geom_point(size = 2) +
      geom_polygon(data = ellipse , fill = "orange" , color = "orange" , alpha = 0.5)+
      geom_point(aes(data.center[1] , data.center[2]) , size = 5 , color = "blue") +
      geom_text( aes(label = row.names(data)) , hjust = 1 , vjust = -1.5 ,size = 2.5 ) +
      ylab("V_Controle") + xlab("OWA")
    
    # Renderiza o gráfico no plotOutput
    output$scatterplot <- renderPlot({
      figure
    })
    
    data.cov <- data.cov + diag(nrow(data.cov)) * 1e-6
    
    # Etapa 4
    # Finding distances
    distances <- mahalanobis(x = data , center = data.center , cov = data.cov)
    
    # Cutoff value for ditances from Chi-Sqaure Dist.
    # with p = 0.95 df = 2 which in ncol(air)
    cutoff <- qchisq(p = 0.95 , df = ncol(data))
    
    # Display observation whose distance greater than cutoff value
    data[distances > cutoff ,]
    
    outliers = nrow(data[distances > cutoff ,])
    outliers = outliers / quantidade_de_elementos_coluna
    
    # Outliers graph
    # Gráfico outliers
    output$histograma <- renderPlot({
      histograma()
    })
    
    output$correlation_plot <- renderPlot({
      correlation_plot()
    })
    
    correlation_plot <- reactive({
      x_values <- unlist(somas_colunas)
      y_values <- unlist(lista_variavel_controle)
      
      if (length(x_values) != length(y_values)) {
        return(NULL)
      }
      
      data_df <- data.frame(OWA = as.numeric(x_values), V_Controle = as.numeric(y_values))
      
      ggscatter(
        data = data_df,
        x = "OWA",
        y = "V_Controle",
        add = "reg.line",
        conf.int = TRUE,
        cor.coef = TRUE,
        cor.method = "pearson",
        title = "Correlation Plot"
      )
    })
    
    lista_correlacao_media <- list()
    
    for (i in 1:nrow(matriz_transposta)) {
      lista_linha <- as.list(matriz_transposta[i, ])
      lista <- as.numeric(lista_linha)
      correlacao <- cor(lista, somas_colunas)
      correlacao <- correlacao^2
      lista_correlacao_media[[i]] <- correlacao
    }
    
    lista_correlacao_media <- as.numeric(lista_correlacao_media)
    media <- mean(lista_correlacao_media)
    
    correlacao_resultado = round(correlacao_resultado, 3)
    media = round(media, 3)
    outliers = round(outliers, 3)
    media_diferenca = abs(round(media_diferenca, 3))
    entropy = round(entropy, 3)
    media_soma = round(media_soma, 3)
    peso = round(peso, 3)
    
    lista_resultados <- list(Explanatory_power = correlacao_resultado, Informational_power = media, Ratio_of_atypical_measurements = outliers, Uncertainty = media_diferenca, Discriminating_power = entropy, Average_scores = media_soma, Orness_degree = peso)
    
    writeData(wb, sheet = "OWA", "Explanatory Power: ", startRow = num_linhas_owa_media + 2, startCol = 1)
    writeData(wb, sheet = "OWA", correlacao_resultado, startRow = num_linhas_owa_media + 2, startCol = 2)
    
    writeData(wb, sheet = "OWA", "Informational Power: ", startRow = num_linhas_owa_media + 3, startCol = 1)
    writeData(wb, sheet = "OWA", media, startRow = num_linhas_owa_media + 3, startCol = 2)
    
    writeData(wb, sheet = "OWA", "Ratio_of_atypical_measurements: ", startRow = num_linhas_owa_media + 4, startCol = 1)
    writeData(wb, sheet = "OWA", outliers, startRow = num_linhas_owa_media + 4, startCol = 2)
    
    writeData(wb, sheet = "OWA", "Uncertainty: ", startRow = num_linhas_owa_media + 5, startCol = 1)
    writeData(wb, sheet = "OWA", media_diferenca, startRow = num_linhas_owa_media + 5, startCol = 2)
    
    writeData(wb, sheet = "OWA", "Discriminating_power: ", startRow = num_linhas_owa_media + 6, startCol = 1)
    writeData(wb, sheet = "OWA", entropy, startRow = num_linhas_owa_media + 6, startCol = 2)
    
    writeData(wb, sheet = "OWA", "Average_scores: ", startRow = num_linhas_owa_media + 7, startCol = 1)
    writeData(wb, sheet = "OWA", media_soma, startRow = num_linhas_owa_media + 7, startCol = 2)
    
    writeData(wb, sheet = "OWA", "Orness_degree: ", startRow = num_linhas_owa_media + 8, startCol = 1)
    writeData(wb, sheet = "OWA", peso, startRow = num_linhas_owa_media + 8, startCol = 2)
    
    
    saveWorkbook(wb, file = nome_do_arquivo)
    
    # Graphics
    # Gráficos
    # Generates Graph 1 - Correlation
    # Gera o Gráfico 1 - Correlação
    generate_correlation_plot <- function(data) {
      chart.Correlation(data, histogram = TRUE, pch = 19)
    }
    
    output$correlationPlot <- renderPlot({
      generate_correlation_plot(planilha_grafico1)
    })
    
    # Generates graph 4 - Uncertainty
    # Gera o gráfico 4 - Incerteza
    output$graficoCandlestick <- renderPlotly({
      diferenca = somas_colunas_fixo - somas_colunas
      
      p <- plot_ly(type="candlestick", open = ordem_eq_resultado_fixo, close = ordem_eq_resultado, high = ordem_eq_resultado_fixo, low = ordem_eq_resultado)
      p <- p %>% layout(width = 800, height = 600)
      
      p
    })
    
    # Generates graph 5 - Histogram - Distribution
    # Gera o gráfico 5 - Histograma - Distribuição
    histograma <- reactive({
      
      chaves <- as.list(chaves)
      chaves <- as.numeric(chaves)
      
      ggplot(data.frame(x = chaves), aes(x)) + geom_histogram(bins = 10, fill = "grey", color = "black") + labs(title = "Histogram", x = "Value", y = "Frequency")
    })
    
    # Download chart 1
    # Faz o download do gráfico 1
    output$download_plot1 <- downloadHandler(
      filename = function() {
        "correlation_plot.png"
      },
      content = function(file) {
        png(file)
        generate_correlation_plot(my_data)
        dev.off()
      }
    )
    
    # Download graph 2
    # Faz o download do gráfico 2
    output$download_plot2 <- downloadHandler(
      filename = function() {
        "mahalanobis.png"
      },
      content = function(file) {
        ggsave(file, figure)  # Salva a figura usando ggsave
      }
    )
    
    # Download chart 3
    # Faz o download do gráfico 3
    output$download_plot3 <- downloadHandler(
      filename = function() {
        "correlacao.png"
      },
      content = function(file) {
        ggsave(file, correlation_plot())  # Salva o gráfico usando ggsave
      }
    )
    
    # Download chart 4
    # Faz o download do gráfico 4
    output$download_plot4 <- downloadHandler(
      filename = function() {
        "uncertanty.png"
      },
      content = function(file) {
        png(file)
        dev.off()
      }
    )
    
    # Download chart 5
    # Faz o download do gráfico 5
    output$download_plot5 <- downloadHandler(
      filename = function() {
        "histogram.png"
      },
      content = function(file) {
        ggsave(file, histograma()) 
      }
    )
    
    return(lista_resultados)
  }
  
  output$selected_output <- renderPrint({
    selected_headers <- input$selected_headers
    selected_header <- input$selected_header
    selected_pn <- input$positive_negative
    selected_slider <- input$custom_slider
    
    output_text <- character(0)
    
    output_text <- c(output_text, paste("Escolha:", selected_pn))
    
    output_text <- c(output_text, paste("Slider:", selected_slider))
    
    output_text
    
    if (input$calcular_owa > 0) {
      list_result <- calcular_owa_function()
      list_text <- list_result
      list_text
    }
    
  })
  observeEvent(input$download_button, {
    
    data_de_hoje <- Sys.Date()
    data_formatada <- format(data_de_hoje, "%d-%m-%Y")
    nome_do_arquivo <- paste0("Passo a passo ",data_formatada,".xlsx")
    file_name <- nome_do_arquivo
    
    output$download_xlsx <- downloadHandler(
      filename = function() {
        file_name
      },
      content = function(file) {
        file.copy(file_name, file)
      }
    )
    
    showModal(modalDialog(
      title = "XLSX file download",
      "Click the button below to download the file with OWA calculations and processes:",
      downloadLink("download_xlsx", "Baixar o arquivo .XLSX"),
      footer = actionButton("fechar_modal", "Fechar")
    ))
  })
  observeEvent(input$fechar_modal, {
    removeModal()
  })
}

shinyApp(ui, server)
