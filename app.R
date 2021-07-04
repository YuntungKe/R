library(shiny)
library(shinyBS)
library(shinyjs)
library(shinythemes)
library(tidyverse)
library(readxl)

path = "listed_unadjusted.xlsx"
stock <- map_dfr(excel_sheets(path), 
                 ~ read_xlsx(path, sheet = .x)) #合併所有分頁
names(stock) <- c("date", "code", "name", "industry", "close", "return", "volume")
stock$return <- stock$return * 0.01

#篩選日成交量大等於一千之個股
unliquid <- stock[stock$volume <1, colnames(stock)== "name"]
cleaned <- stock[!stock$name %in% unliquid$name, ]

#移除空股票代號、資料筆數不足股票
code_valid <- as.numeric(intersect(x = c(1101:9958), y = cleaned$"code"))

R <- function(code){
  r <- cleaned[which(cleaned$"code" == code), colnames(cleaned)=="return"]
  return(r)}

List_complete = list()
for (v in c(code_valid)) {
  if(nrow(R(v)) == 728){
    complete <- as.numeric(v)
    List_complete[[length(List_complete)+1]] = complete
  }
}  
List_complete <- as.data.frame(List_complete) #代號存在且資料筆數完整
clean <- cleaned[which(cleaned$"code" %in% List_complete),]

load("ifrsProfitability.rda")
profitIndex  <- profitability[profitability$"年月" >= 201603,] #近五年獲利能力
profitIndex <- profitIndex[profitIndex$公司代號 %in% List_complete,]

#ROE近4季>5%
ROE <- data.frame(code = profitIndex[profitIndex$"年月" >= 201912 & profitIndex$ROE > 0.05, 1]) 
ROE <- (ROE %>% count(code) %>% filter(n == 4))[1]
wellROE <- clean[which(clean$"code" %in% ROE$"code"),]
#毛利率近五年>10%
grossProf <- data.frame(code = profitIndex[profitIndex$ROE > 0.1, 1])
grossProf <- (grossProf %>% count(code) %>% filter(n == max(n)))[1]
wellGP <- clean[which(clean$"code" %in% grossProf$"code"),]
#wellROEnGP <- intersect(ROE, grossProf) #兩者皆符合者恰與毛利率名單相同

List_industry <- list("水泥工業" = "01", "食品工業" = "02", "塑膠工業" = "03", "紡織纖維" = "04", "電機機械" = "05",
                      "電器電纜" = "06", "化學工業" = "21", "生技醫療" = "22", "玻璃陶磁" = "08", "造紙工業" = "09",
                      "鋼鐵工業" = "10", "橡膠工業" = "11", "汽車工業" = "12", "半導體業" = "24", "電腦及周邊設備業" = "25",
                      "光電業" = "26", "電子零組件業" = "28", "電子通路業" = "29", "其他電子業" = "31", "通信網路業" = "27",
                      "資訊服務業" = "30", "營造建材" = "14", "航運業" = "15", "觀光業" = "16", "金融業" = "17",
                      "貿易百貨" = "18", "油電燃氣業" = "23", "其他" = "20")

jsResetCode <- "shinyjs.reset = function() {history.go(0)}"

###ui######################################################
ui <- fluidPage(theme = shinytheme("flatly"),
                
                navbarPage(
                  "Stock Risk Management",
                  
                  tabPanel("持股健檢",
                           sidebarPanel(
                             h3("輸入："),
                             textInput("txt1", p("個股代碼",popify(a(icon("fas fa-external-link-alt"), href = "https://tw.stock.yahoo.com/h/kimosel.php"),
                                                               title = "代碼查詢",
                                                               content = paste0("點選 ", icon = icon("fas fa-external-link-alt"), "前往Yahoo!奇摩股市 - 股票代號查詢"),
                                                               placement = "right")
                             ), 
                             placeholder = "以「空白鍵」隔開 (ex.1101 2330 ...)"),
                             textInput("txt2", "個股張數", placeholder = "以「空白鍵」隔開 (ex.3 1 ...)"),
                             actionButton("submitbutton", "Submit", class = "btn btn-primary"),
                             
                             useShinyjs(),                                           
                             extendShinyjs(text = jsResetCode, functions = "reset"), # Add the js code to the page
                             popify(actionButton("reset_button", "Reset", icon = icon("redo")),
                                    title = "重新整理", 
                                    content = "按此以重整頁面 或使用快捷鍵Ctrl+R",
                                    placement = "right")
                           ), # sidebarPanel
                           
                           mainPanel(
                             h1("持股健檢"),
                             
                             h4("您的持股"),
                             verbatimTextOutput("txtresult"),
                             verbatimTextOutput("txtresult2"),
                             h5("若上方出現紅字警告，請再次檢查輸入格式是否以「空白鍵」隔開"),
                             
                             tags$label(h3('Status/Output')), 
                             verbatimTextOutput("txtresult3"),
                             verbatimTextOutput("observe"),
                             #span(textOutput("observe"), style="color:red"),
                             
                             plotOutput("plot"),
                             span(textOutput("descrip")),
                             tableOutput('tabledata'),
                             textOutput("info"),
                             tableOutput('tabledata2'),
                             imageOutput("img")
                           ), #mainPanel
                           
                           span(htmlOutput("observe2"), align = "center"),
                           bsPopover("observe2", 
                                     title = "重整頁面",
                                     content = paste0("上方 ", icon = icon("redo"), "Reset 鈕"),
                                     placement = "bottom"),
                           
                           span(".", style="color:white"),
                           hr(),
                           h5("Sources:"),
                           h6(
                             p(a("TEJ", href = "https://www.tej.com.tw/"),"台灣經濟新報資料庫 未調整股價資料")),
                           h6(p("歡迎提供建議與回饋至 yt.ke.data@gmail.com")),
                           h6("2021 by Yuntung Ke. Built with",
                              img(src = "https://www.rstudio.com/wp-content/uploads/2014/04/shiny.png", height = "30px"),
                              "by",
                              img(src = "https://www.rstudio.com/wp-content/uploads/2014/07/RStudio-Logo-Blue-Gray.png", height = "30px"),
                              ".")
                  ), #navbar 1
                  
                  tabPanel("投資組合建構",
                           sidebarPanel(
                             h3("輸入："),
                             sliderInput("riskLevel", "風險等級(請參考右表):",
                                         min = 1, max = 6,
                                         value = 1),
                             
                             selectInput(
                               "industry_select1",
                               label = "請挑選任一產業",
                               choices = List_industry,
                               selected = "01"
                             ),
                             
                             selectInput(
                               "industry_select2",
                               label = "請挑選任一產業",
                               choices = List_industry,
                               selected = "01"
                             ),
                             
                             checkboxGroupInput(
                               "prof",
                               "獲利能力",
                               choices = list("無限制" = 1, "近4季ROE>5%" = 2,
                                              "近5年毛利率>10%" = 3, "同時符合前二者" = 4),
                               selected = NULL,
                               inline = FALSE
                             ),
                             
                             textInput("txt4", "投資金額", placeholder = "請輸入欲投入金額"),
                             actionButton("submitbutton2", "Submit", class = "btn btn-primary")
                           ), # sidebarPanel
                           
                           mainPanel(
                             h1("投組建構"),
                             
                             imageOutput("img4"),
                             
                             tags$label(h3('Status/Output')),
                             verbatimTextOutput("txtresult11"),
                             verbatimTextOutput("observe3"),
                             plotOutput("plot2"),
                             
                             span(textOutput("tell")), 
                             tableOutput('tabledata3'),
                             span(textOutput("info2")),
                             tableOutput('tabledata4'),
                             imageOutput("img6"),
                           ), #mainPanel
                           
                           span(htmlOutput("observed"), align = "center"),
                           
                           span(".", style="color:white"),
                           hr(),
                           h5("Sources:"),
                           h6(
                             p(a("TEJ", href = "https://www.tej.com.tw/"),"台灣經濟新報資料庫 未調整股價資料")),
                           h6(p("歡迎提供建議與回饋至 yt.ke.data@gmail.com")),
                           h6("2021 by Yuntung Ke. Built with",
                              img(src = "https://www.rstudio.com/wp-content/uploads/2014/04/shiny.png", height = "30px"),
                              "by",
                              img(src = "https://www.rstudio.com/wp-content/uploads/2014/07/RStudio-Logo-Blue-Gray.png", height = "30px"),
                              ".")
                  ), #navbar 2
                  
                  tabPanel("Data & Methodology", fluid = TRUE,
                           fluidRow(
                             column(6,
                                    h4(strong("關於資料")),
                                    h5(strong("___________________________________________________________________")),
                                    br(),
                                    h5(p(strong("持股健檢"),"程式採用", a("TEJ", href = "https://www.tej.com.tw/"),"台灣經濟新報資料庫2018-2020年未調整股價(日)資料，")),
                                    h5(p("並篩選日成交量大於1千(張)之上市個股。")),
                                    br(),
                                    br(),
                                    br(),
                                    h5(p("▽VaR示意圖")),
                                    imageOutput("img2"),
                             ),
                             column(6,
                                    h4(strong("關於VaR")),
                                    h5(strong("___________________________________________________________________")),
                                    h5(p(a("風險價值", href = "https://reurl.cc/kZ1dpr"),"（Value at Risk，縮寫VaR），衡量市場風險的一種方法。")),
                                    
                                    h5(p("其意義為在特定期間及特定機率下，持有單一資產或資產的投資組合，")),
                                    h5(p("因市場上經濟變數之變動，預期該組合可能產生的最大損失。是風險管")),
                                    h5(p("理中應用廣泛、研究活躍的風險定量分析方法之一。")),
                                    br(),
                                    h5(p("Hull與White(1998)提出對VaR的定義為：「有100(1-α)%的信心在未來")),
                                    h5(p("N天內的最大損失不會超過V元」，其中V即為風險值。")),
                                    
                                    br(),
                                    h5(strong("常見的VaR估算方法：")),
                                    h5(p("．分析法(Analytic Method)")),
                                    h5(p("．法歷史模擬法(Historical Method)")),
                                    h5(p("．蒙地卡羅模擬Monte Carlo Simulation Method)")),
                                    
                                    br(),
                                    h5(strong("▽ 持股健檢"),"程式採用",strong("分析法"),"，公式如下"),
                                    imageOutput("img3")
                             ),
                             column(10,
                                    h4(strong("關於風險等級")),
                                    h5(strong("_____________________________________________________________________________")),
                                    h5("風險等級量表參考",strong("歐洲證券及市場管理局"),a("(ESMA)", href = "https://www.esma.europa.eu/"),("所編制之ESMA volatility bands，")),
                                    h5(p("將其「週波動」區間以對數計算轉換為「日波動」區間，並重新分割及定義風險等級。")),
                                    
                                    br(),
                                    h5(strong("▽ ESMA volatility bands")),
                                    imageOutput("img5")
                             )
                           ),
                           
                           span(".", style="color:white"),
                           hr(),
                           h5("Sources:"),
                           h6(
                             p("VaR說明截取自",
                               a("元富證券",
                                 href = "https://www.masterlink.com.tw/about/riskmanagment/manage/market.aspx"),
                               "&", a("BME CLEARING",
                                      href = "https://www.bmeclearing.es/ing/Risk-Management/IM-Calculation-model-HVAR"))),
                           h6(p("歡迎提供建議與回饋至 yt.ke.data@gmail.com")),
                           
                           h6("2021 by Yuntung Ke. Built with",
                              img(src = "https://www.rstudio.com/wp-content/uploads/2014/04/shiny.png", height = "30px"),
                              "by",
                              img(src = "https://www.rstudio.com/wp-content/uploads/2014/07/RStudio-Logo-Blue-Gray.png", height = "30px"),
                              ".")
                           
                  ) #navbar 3
                  
                ) #navbarPage
)


##server###################################################
server <- function(input, output, session) {
  
  observeEvent(input$reset_button, {js$reset()})
  
  mydata1 <- reactive({
    mystock <-scan(text=input$txt1,quiet=TRUE,sep="")
    mystock
  })
  
  mydata2 <- reactive({
    myquant <-scan(text=input$txt2,quiet=TRUE,sep="")
    myquant
  })
  
  ##output 1####
  output$txtresult <- renderText({
    mystock <-  mydata1()
    myquant <-  mydata2()
    
    if(length(mystock) == 0){
      code_quant <- as.character("請輸入持股")
    }else{
      code_quant = NULL
      for(i in 1:length(mystock)){
        if(mystock[i] %in% List_complete){
          code_quant[i] <- paste(c(mystock[i], as.character(cleaned[cleaned$code == mystock[i],colnames(cleaned)== "name"][1,]),
                                   myquant[i], "張" ,'\n'), collapse = " ")
        }else{
          code_quant <- as.character("抱歉，個股代碼輸入錯誤 或 該股資料筆數不足無法計算")
        }
      }
    }
    
    code_quant
  })
  
  #####投組數據計算
  #投組價值
  mydata3 <- reactive({
    mystock <-  mydata1()
    myquant <-  mydata2()
    value = NULL
    for (i in 1:length(mystock)) {
      value[i] <- sum((cleaned[which(cleaned$"code"== mystock[i] & cleaned$"date" == max(cleaned$"date")),5])*myquant[i])
    }
    value
  })
  
  #投組權重
  mydata4 <- reactive({
    mystock <- mydata1()
    myquant <- mydata2()
    value <- mydata3()
    weight = NULL
    for (i in 1:length(mystock)) {
      weight[i] <- value[i]/sum(value)
    } 
    weight
  }) 
  
  #投組平均日報酬、日標準差
  mydata5 <- reactive({
    mystock <- mydata1()
    myquant <- mydata2()
    value <- mydata3()
    weight <- mydata4()
    weightReturn = NULL
    for (i in 1:length(mystock)) {
      weightReturn[i] <- weight[i] * cleaned[which(cleaned$"code"== mystock[i]),6]
    }
    weightReturn <- as.data.frame(weightReturn) #投組日報酬(已乘權重)
    weightReturn
  }) 
  
  ##output 2####
  output$txtresult2 <- renderText({
    mystock <- mydata1()
    myquant <- mydata2()
    value <- mydata3()
    weight <- mydata4()
    weightReturn <- mydata5()
    
    avgReturn <- mean(rowSums(weightReturn)) #投組"平均"日報酬
    sdReturn <- sd(rowSums(weightReturn)) #投組"平均"日標準差
    
    #VaR分析法= 報酬率- Zα* σ   # Zα(95%)= 1.65
    portVaR <- avgReturn - 1.65 * sdReturn
    pVaR <- signif(portVaR,2)
    percent_portVaR <- -signif(portVaR, 4) * 100
    
    showVaR <- paste("在95%信心水準下，此投組最大日損失為", percent_portVaR, "% 投組價值",
                     "=", sum(value) * 1000 * -pVaR, "元")
    
    showVaR
  })
  
  #########
  datasetInput <- reactive({  
    mystock <- mydata1()
    myquant <- mydata2()
    value <- mydata3()
    weight <- mydata4()
    weightReturn <- mydata5()
    
    avgReturn <- mean(rowSums(weightReturn)) #投組"平均"日報酬
    sdReturn <- sd(rowSums(weightReturn)) #投組"平均"日標準差
    
    #VaR分析法= 報酬率- Zα* σ   # Zα(95%)= 1.65
    portVaR <- avgReturn - 1.65 * sdReturn
    percent_portVaR <- -signif(portVaR,4) * 100
    
    
    #####尋找避險股
    #尋找與投組最小相關係數
    List_return = list()
    for (w in c(List_complete)) {
      everyRet <- R(w)
      List_return[[length(List_return)+1]] = everyRet
    }
    List_return <- as.data.frame(List_return)
    colnames(List_return) <- c(List_complete) #資料完整個股之報酬
    
    Corr = matrix(NA_integer_, nrow = length(mystock), ncol = ncol(List_return))
    for (i in 1:length(mystock)){
      for(w in 1:ncol(List_return)){
        Corr[i,w] <- cor(cleaned[which(cleaned$"code"== mystock[i]),6], List_return[w])
      }
    }
    Corr <- as.data.frame(Corr, row.names = c(mystock)) #相關係數表
    colnames(Corr) <- c(List_complete)
    
    #找最小相關
    minCorr <- data.frame(sum_corr = apply(Corr, 2, sum)) 
    minCorr <- arrange(minCorr, sum_corr)
    top10_sum_corr <- row.names(minCorr)[1:10] #取最小相關前10名個股
    
    
    #####避險投組數據計算
    #避險股權應購張數
    hedge_num = NULL
    for (j in 1:length(top10_sum_corr)) {
      hedge_num[j] <- as.numeric(sum(value)*0.25 / #預設權重為20%(可改)
                                   cleaned[which(cleaned$"code"== top10_sum_corr[j] & cleaned$"date" == max(cleaned$"date")),5]) 
    }
    hedge_num <- signif(hedge_num,2)
    
    #投組價值(避險)
    hedge_value = NULL
    hedge_weight = matrix(NA_integer_, nrow = length(mystock), ncol = length(top10_sum_corr))
    for (j in 1:length(top10_sum_corr)) {
      hedge_value[j] <- sum(value) + 
        sum((cleaned[which(cleaned$"code"== top10_sum_corr[j] & cleaned$"date" == max(cleaned$"date")),5])* hedge_num[j])
    }
    
    #投組權重(避險)
    for (i in 1:length(mystock)) {
      for (j in 1:length(top10_sum_corr)) {
        hedge_weight[i,j] <- value[i]/hedge_value[j]
      }
    }
    hedge_weight <- data.frame(hedge_weight, row.names = c(mystock))
    colnames(hedge_weight) <- c(top10_sum_corr)
    add_weight <- 1- apply(hedge_weight, 2, sum)
    hedge_weight <- rbind(hedge_weight,add_weight)
    
    #投組平均日報酬、日標準差(避險)
    #避險股報酬
    top10_return = matrix(NA_integer_, nrow = length(top10_sum_corr), ncol = 1)
    for (j in 1:length(top10_sum_corr)) {
      top10_return[j] <- as.data.frame(cleaned[which(cleaned$"code"== top10_sum_corr[j]),6])
    }
    top10_return <- as.data.frame(top10_return)
    colnames(top10_return) <- c(top10_sum_corr)
    
    #原持股*新權重hedge_weight
    mystock_hedgeWeight = list()
    for (j in 1:length(top10_sum_corr)) {
      for (i in 1:length(mystock)) {
        a <- as.list(R(mystock[i])*hedge_weight[i,j])
        mystock_hedgeWeight[[length(mystock_hedgeWeight)+1]] = a
      }
    }
    mystock_hedgeWeight <- as.data.frame(mystock_hedgeWeight)
    colnames(mystock_hedgeWeight) <- rep(c(mystock),length(top10_return))
    
    #避險股*權重hedge_weight
    top10_hedge_weight = matrix(NA_integer_, nrow = length(top10_sum_corr), ncol = 1)
    for (j in 1:length(top10_sum_corr)) {
      top10_hedge_weight[j] <- as.data.frame(hedge_weight[length(mystock)+1,j] * cleaned[which(cleaned$"code"== top10_sum_corr[j]),6])  
    }
    top10_hedge_weight <- as.data.frame(top10_hedge_weight)
    colnames(top10_hedge_weight) <- c(top10_sum_corr)
    
    #避險後投組
    hedge_return = NULL
    for (j in 1:length(top10_sum_corr)) {
      for(h in seq(from = 1, to = 10*length(mystock), by = length(mystock)))
        hedge_return[j] <- top10_hedge_weight[j] + apply(mystock_hedgeWeight[h:(h+length(mystock)-1)], 1, sum)
    }
    hedge_return <-  as.data.frame(hedge_return)
    colnames(hedge_return) <- c(top10_sum_corr)
    
    hedge_ret_sd <- data.frame(平均日報酬率 = apply(hedge_return, 2, mean),#投組平均日報酬(避險)
                                     平均日標準差 = apply(hedge_return, 2, sd))#投組平均日標準差(避險)
    hedge_portVaR <- data.frame(VaR = hedge_ret_sd$"平均日報酬率" - 1.65 * hedge_ret_sd$"平均日標準差") #投組VaR(避險)
    hedge_port <- data.frame(hedge_portVaR, hedge_ret_sd[1], 應購張數 = hedge_num)#避險結果
    
    #篩選避險後有改善風險之個股
    well_hedge <- hedge_port[which(hedge_port$VaR > portVaR),]
    well_hedge
  })
  
  datasetInput4 <- reactive({
    well_hedge <- datasetInput()
    
    num_hedge  <- nrow(well_hedge)
    num_hedge
  })
  
  datasetInput5 <- reactive({
    weightReturn <- mydata5()
    avgReturn <- mean(rowSums(weightReturn)) #投組"平均"日報酬
    sdReturn <- sd(rowSums(weightReturn)) #投組"平均"日標準差
    #VaR分析法= 報酬率- Zα* σ   # Zα(95%)= 1.65
    portVaR <- avgReturn - 1.65 * sdReturn
    percent_portVaR <- -signif(portVaR,4) * 100
    well_hedge <- datasetInput()
    
    well_hedge[nrow(well_hedge)+1,] <- data.frame(origin = list(portVaR,avgReturn,c("")), row.names = c("原投組"))#比較
    well_hedge <- well_hedge[order(row.names(well_hedge)),]
    
    well_hedge$VaR變化 <- c("+")
    well_hedge$報酬率變化 <- ifelse(well_hedge$平均日報酬率 > avgReturn, "+", "-")
    well_hedge$加入個股 <- c(row.names(well_hedge))
    well_hedge <- well_hedge[c(6,1,4,2,5,3)]; well_hedge <- well_hedge[(c(nrow(well_hedge),1:nrow(well_hedge)-1)),]
    well_hedge[1,c(3,5)] = c("")
    
    well_hedge
  })
  
  #########    
  datasetInput2 <- reactive({
    well_hedge <- datasetInput5()
    
    names(stock) <- c("日期", "個股代碼", "簡稱", "產業別", "收盤價", "報酬率", "千成交量")
    df_well_hedge <- stock[which(stock$'個股代碼' %in% rownames(well_hedge) & stock$'日期' == max(stock$'日期')),c(2:5)]
    
    df_well_hedge
  })
  
  datasetInput3 <- reactive({ 
    well_hedge <- datasetInput5()
    
    ret_sd <- data.frame((well_hedge$VaR) *-100,(well_hedge$"平均日報酬率")*100)
    colnames(ret_sd) <- c("日最大損失","日報酬率")
    rownames(ret_sd) <- row.names(well_hedge)
    ret_sd
  })
  
  ##output 3####  
  output$txtresult3 <- renderPrint({
    if (input$submitbutton>0) { 
      cat("Calculation complete.") 
    } else {
      cat("Server is ready for calculation.")
    }
  })
  
  ##output 4####
  output$observe <- renderPrint({
    if (input$submitbutton>0) { 
      if(datasetInput4 () == 0){
        stop("您的風險管理近乎完美!已找不到可加入之避險股!!再見")
      }else {
        cat("↓↓以下圖表為加入20%投組價值之個股的結果↓↓")
      }
    }
  })
  
  output$observe2 <- renderPrint({
    if (input$submitbutton>0) { 
      if(datasetInput4 () == 0){
        stop()
      }else {
        cat("- 感謝您的使用！如欲再降低投組風險，請","<font color=\"#D2B48C\"><u>","重整頁面","</u></font>","並加入避險股以獲建議 -")
      }
    }
  })
  
  output$descrip <- renderPrint({
    if (input$submitbutton>0) { 
      if(datasetInput4 () == 0){
        stop()
      }else {
        cat("+ 表加入該股後改善, - 表變差")
      }
    }
  })
  
  output$info <- renderPrint({
    if (input$submitbutton>0) { 
      if(datasetInput4 () == 0){
        stop()
      }else {
        cat("個股資訊：")
      }
    }
  })
  
  ##plot########
  output$plot <- renderPlot({
    if (input$submitbutton>0) { 
      ret_sd <- datasetInput3()
      weightReturn <- mydata5()
      avgReturn <- mean(rowSums(weightReturn)) #投組"平均"日報酬
      sdReturn <- sd(rowSums(weightReturn)) #投組"平均"日標準差
      
      #VaR分析法= 報酬率- Zα* σ   # Zα(95%)= 1.65
      portVaR <- avgReturn - 1.65 * sdReturn
      percent_portVaR <- -signif(portVaR,4) * 100
      
      #作圖
      plot(ret_sd,
           main = "投資組合 日最大損失-日報酬 圖",
           xlab = "日最大損失(%投組價值)",
           xlim = c(min(ret_sd$日最大損失)-0.01,max(ret_sd$日最大損失)+0.12),
           ylab = "日報酬率(%)")
      points(ret_sd[1,], pch = 16)
      text(ret_sd,labels = row.names(ret_sd),pos=4) 
      abline(h = avgReturn*100, v = percent_portVaR, lty = 2)
      text(x = (min(ret_sd$日最大損失) + percent_portVaR)/2,
           y =  (max(ret_sd$日報酬率) + avgReturn*100)/2, 
           labels = "風險下降，報酬上升",
           col = "lightsteelblue1")
      text(x = ( min(ret_sd$日最大損失) + percent_portVaR)/2, 
           y =  (min(ret_sd$日報酬率) + avgReturn*100)/2, 
           labels = "風險下降，報酬下降",
           col = "gray80")
      grid()
    } 
  })
  
  ##tb 1########
  output$tabledata <- renderTable({
    if (input$submitbutton>0) { 
      isolate(datasetInput5()) 
    } 
  },digits = 6)
  
  ##tb 2########
  output$tabledata2 <- renderTable({
    if (input$submitbutton>0) { 
      isolate(datasetInput2()) 
    } 
  })
  
  observeEvent(input$submitbutton, {
    pic <- list(src="industry.png", width="100%", alt = "錯誤")
    output$img <- renderImage({
      pic
    }, deleteFile=FALSE)
  })  
  
  output$img2 <- renderImage({
    list(src="dis_var.png", width="90%", alt = "錯誤")
  }, deleteFile=FALSE)
  
  output$img3 <- renderImage({
    list(src="formula.png", width="50%", alt = "錯誤")
  }, deleteFile=FALSE)
  
  ##二、投組建構#########################################################
  #風險-日標準差
  data1 <- reactive({
    y <- input$riskLevel
    if(y == 1){
      myVolLow = 0.0061051
      myVolUp = 0.007623
    }else if(y == 2){
      myVolLow = 0.0076231
      myVolUp = 0.010114
    }else if(y == 3){
      myVolLow = 0.0101141
      myVolUp = 0.014883
    }else if(y == 4){
      myVolLow = 0.0148831
      myVolUp = 0.020518
    }else if(y == 5){
      myVolLow = 0.0205181
      myVolUp = 0.027129
    }else if(y == 6){
      myVolLow = 0.0271291
      myVolUp = 1
    }
    
    myVolLow
  })
  
  data11 <- reactive({
    y <- input$riskLevel
    if(y == 1){
      myVolLow = 0.0061051
      myVolUp = 0.007623
    }else if(y == 2){
      myVolLow = 0.0076231
      myVolUp = 0.010114
    }else if(y == 3){
      myVolLow = 0.0101141
      myVolUp = 0.014883
    }else if(y == 4){
      myVolLow = 0.0148831
      myVolUp = 0.020518
    }else if(y == 5){
      myVolLow = 0.0205181
      myVolUp = 0.027129
    }else if(y == 6){
      myVolLow = 0.0271291
      myVolUp = 1
    }
    
    myVolUp
  })
  
  data12 <- reactive({
    profitSelect <- input$prof
    
    if(profitSelect == 1){
      clean <- clean
    }else if(profitSelect == 2){
      clean <- wellROE
    }else if(profitSelect == 3 | 4){
      clean <- wellGP
    }
    
    clean
  })
  
  data2 <- reactive({
    clean <- data12()
    
    industry1 <- input$industry_select1
    indSelect1 <- data.frame(unique(clean[which(clean$"industry" == industry1), 2])) 
    indSelect1
  })
  
  data13 <- reactive({
    indSelect1 <- data2()
    
    num_indSelect1  <- nrow(indSelect1)
    num_indSelect1
  })
  
  data3 <- reactive({
    clean <- data12()
    
    industry2 <- input$industry_select2
    indSelect2 <- data.frame(unique(clean[which(clean$"industry" == industry2), 2])) 
    indSelect2
  })
  
  data14 <- reactive({
    indSelect2 <- data3()
    
    num_indSelect2  <- nrow(indSelect2)
    num_indSelect2
  })
  
  data4 <- reactive({
    clean <- data12()
    indSelect1 <- data2()
    indSelect2 <- data3()
    
    cbn <- expand.grid(indSelect1[1:nrow(indSelect1),], 
                       indSelect2[1:nrow(indSelect2),])
    
    #投組平均日報酬、日標準差
    weightReturn2 = NULL
    for (i in 1:nrow(cbn)) {
      weightReturn2[i] <- 0.5 * (clean[which(clean$"code"== cbn[i,1]),6]+
                                       clean[which(clean$"code"== cbn[i,2]),6])
    }
    weightReturn2 <- as.data.frame(weightReturn2) #投組日報酬(已乘權重)
    colnames(weightReturn2) <- c(row.names(cbn)) #組合對照cbn
    
    avgReturn2 = NULL
    for (i in 1:nrow(cbn)) {
      avgReturn2[i] <- mean(rowSums(weightReturn2[i]))
    } #投組"平均"日報酬
    
    sdReturn2 = NULL
    for (i in 1:nrow(cbn)) {
      sdReturn2[i] <- sd(rowSums(weightReturn2[i]))
    } #投組"平均"日標準差
    
    
    myVolLow <- data1()
    myVolUp <- data11()
    
    result <- rbind(sdReturn2,avgReturn2,weightReturn2)
    rownames(result)[rownames(result) == 1] <- "sdReturn"
    rownames(result)[rownames(result) == 2] <- "avgReturn"
    r <- as.data.frame(t(result))
    r <- r[which(myVolLow <= r$sdReturn &  r$sdReturn <= myVolUp & r$avgReturn > 0), c(1,2)]
    namelist <- cbn[row.names(cbn) %in% row.names(r),]
    result <- cbind(namelist,r)
    result <- arrange(result,desc(sdReturn))
    
    result
  })
  
  data5 <- reactive({
    valInvest <-scan(text=input$txt4,quiet=TRUE)
    valInvest
  })
  
  #應購張數
  data6 <- reactive({
    clean <- data12()
    result <- data4()
    valInvest <- data5()
    
    inv_num1 = NULL
    for (j in 1:nrow(result)) {
      inv_num1[j] <- as.numeric(valInvest/1000 * 0.5 / 
                                  clean[which(clean$"code"== result[j,1] & clean$"date" == max(clean$"date")),5])
    }
    inv_num1 <- signif(inv_num1,2)
    
    inv_num1
  })
  
  data7 <- reactive({
    clean <- data12()
    result <- data4()
    valInvest <- data5()
    
    inv_num2 = NULL
    for (j in 1:nrow(result)) {
      inv_num2[j] <- as.numeric(valInvest/1000 * 0.5 / 
                                  clean[which(clean$"code"== result[j,2] & clean$"date" == max(clean$"date")),5])
    }
    inv_num2 <- signif(inv_num2,2)
    
    inv_num2
  })
  
  data8 <- reactive({
    result <- data4()
    valInvest <- data5()
    
    colnames(result) <- c("code","code","sdReturn","avgReturn")
    stock_recmd <- rbind(result[1],result[2])
    stock_recmd <- unique(stock_recmd)
    
    stock_recmd
  })
  
  data9 <- reactive({
    result <- data4()
    valInvest <- data5()
    inv_num1 <- data6()
    inv_num2 <- data7()
    
    result <- data.frame("投組編號" = c(row.names(result)),"個股1" = result[1:nrow(result),1], "應購張數1" = inv_num1,
                         "個股2" = result[1:nrow(result),2], "應購張數2" = inv_num2,
                         "平均日標準差(波動)" = result[1:nrow(result),3], "平均日報酬" = result[1:nrow(result),4])
    
    result
  })
  
  data10 <- reactive({
    clean <- data12()
    stock_recmd <- data8()
    
    names(clean) <- c("日期", "個股代碼", "簡稱", "產業別", "收盤價", "報酬率", "千成交量")
    df_recmd <- clean[which(clean$'個股代碼' %in% t(stock_recmd) & clean$'日期' == max(clean$'日期')),c(2:5)]
    
    df_recmd
  })
  
  ##tb 1########
  output$tabledata3 <- renderTable({
    if (input$submitbutton2 > 0) { 
      if(data13() > 0){
        if(data14() > 0){
          isolate(data9()) 
        }
      }
    } 
  },digits = 6)
  
  ##tb 2########
  output$tabledata4 <- renderTable({
    if (input$submitbutton2 > 0) {
      if(data13() > 0){
        if(data14() > 0){
          isolate(data10()) 
        } 
      }
    }
  },digits = 2)
  
  
  ##plot########
  output$plot2 <- renderPlot({
    if (input$submitbutton2 > 0) { 
      if(data13() > 0){
        if(data14() > 0){
          result <- data9()
          
          port_rcmd <- data.frame(result[6] * 100, result[7] * 100)
          tai <- data.frame("平均日標準差.波動." = 1.035, "平均日報酬" = 0.04889, row.names = c("大盤"))
          port_rcmd <- rbind(port_rcmd,tai)
          
          plot(port_rcmd,
               main = "投資組合 日標準差(波動)-日報酬 圖",
               xlab = "日標準差(%)",
               xlim = c(min(port_rcmd$"平均日標準差.波動.")-0.01,max(port_rcmd$"平均日標準差.波動.")+0.05),
               ylab = "日報酬率(%)",
               ylim = c(min(port_rcmd$"平均日報酬")-0.01,max(port_rcmd$"平均日報酬")+0.01))
          points(port_rcmd[nrow(port_rcmd),], pch = 16)
          text(port_rcmd,labels = row.names(port_rcmd),pos=4) 
          abline(h = 0.04889, v = 1.035, lty = 2)
          text(x = (min(port_rcmd$"平均日標準差.波動.") + 1.035)/2,
               y =  (max(port_rcmd$平均日報酬) + 0.04889)/2, 
               labels = "風險<大盤，報酬>大盤",
               col = "lightsteelblue1")
          legend("topleft",
                 pch = c(1,NA),
                 lty = c(NA,1),
                 legend = c("投組編號")
          )
          grid()
        }
      }
    } 
  })
  
##txt#########
  output$txtresult11 <- renderPrint({
    if (input$submitbutton2 > 0) { 
      cat("Calculation complete.") 
    } else {
      cat("Server is ready for calculation.")
    }
  })
  
  output$observe3 <- renderPrint({
    if (input$submitbutton2 > 0) { 
      if(data13() == 0){
        stop("產業1無符合條件之個股，請重新選擇")
      }else if(data14() == 0){
        stop("產業2無符合條件之個股，請重新選擇")
      }else {
        cat("↓↓下圖表投組為兩個股以1:1權重建構之結果↓↓")
      }
    }
  })

  output$tell <- renderPrint({
    if (input$submitbutton2 > 0) {
      if(data13()>0){
        if(data14()>0){
          cat("*大盤指數平均日標準差為0.01035, 日報酬為0.0004889*")
      }
      }else{cat("")}
    }
  }) 
  
  output$info2 <- renderPrint({
    if (input$submitbutton2 > 0) {
      if(data13()>0){
        if(data14()>0){
          cat("*個股資訊：")
      }
      }
    }else{cat("")}
  }) 
  
  output$observed <- renderPrint({
    if (input$submitbutton2 > 0) {
      cat("- 感謝您的使用 -")
    }
  }) 
  
##img#########
  output$img4 <- renderImage({
    list(src="risk.png", width="100%", alt = "錯誤")
  }, deleteFile=FALSE)
  
  output$img5 <- renderImage({
    list(src="esma.png", width="60%", alt = "錯誤")
  }, deleteFile=FALSE)
  
  observeEvent(input$submitbutton2, {
    pic2 <- list(src="industry.png", width="100%", alt = "錯誤")
    output$img6 <- renderImage({
      pic2
    }, deleteFile=FALSE)
  })  
}

##shinyApp#########
shinyApp(ui, server)