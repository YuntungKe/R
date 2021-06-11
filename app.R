library(shiny)
library(shinythemes)
library(rsconnect)
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


###ui######################################################
ui <- fluidPage(theme = shinytheme("flatly"),
                navbarPage(
                  
                  "Stock Risk Management",
                  tabPanel("持股健檢",
                           sidebarPanel(
                             h3("輸入："),
                             textInput("txt1", "個股代碼", placeholder = "以逗號隔開(ex.1101,2330,...)"),
                             textInput("txt2", "個股張數", placeholder = "以逗號隔開(ex.2,3,...)"),
                             actionButton("submitbutton", "Submit", class = "btn btn-primary")
                           ), # sidebarPanel
                           
                           mainPanel(
                             h1("持股健檢"),
                             
                             h4("您的持股"),
                             verbatimTextOutput("txtresult"),
                             verbatimTextOutput("txtresult2"),
                             
                             tags$label(h3('Status/Output')), 
                             verbatimTextOutput("txtresult3"),
                             plotOutput("plot"),
                             h5("+ 表加入該股後改善, - 表變差"),
                             tableOutput('tabledata'),
                             h5("個股資訊："),
                             tableOutput('tabledata2'),
                             imageOutput("img")
                             
                           ) #mainPanel
                  ), #navbar 1
                  
                  tabPanel("投資組合推薦", "This panel is intentionally left blank"),
                  
                  tabPanel("Data & Methodology", fluid = TRUE,
                           fluidRow(
                             column(6,
                                    #br(),
                                    h4(strong("關於資料")),
                                    h5(strong("___________________________________________________________________")),
                                    br(),
                                    h5(p(strong("持股健檢"),"程式採用", a("TEJ", href = "https://www.tej.com.tw/"),"台灣經濟新報資料庫2016-2020年未調整股價(日)資料，")),
                                    h5(p("並篩選日成交量大於1千(張)之上市個股。")),
                                    br(),
                                    br(),
                                    br(),
                                    h5(p("▽VaR示意圖")),
                                    imageOutput("img2"),
                                    
                             ),
                             column(6,
                                    #br(),
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
                                    
                             )
                           ),
                           
                           shiny::hr(),
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
                           
                  ) #navbar 2
                  
                ) #navbarPage
)


##server###################################################
server <- function(input, output, session) {
  
  extract <- function(text) {
    text <- gsub(" ", "", text)
    split <- strsplit(text, ",", fixed = FALSE)[[1]]
    as.numeric(split)
  }
  
  mydata1 <- reactive({
    nums1 <- extract(input$txt1)
    mystock <-scan(text=input$txt1,quiet=TRUE,sep=",")
    mystock
  })
  
  mydata2 <- reactive({
    nums2 <- extract(input$txt2)
    myquant <-scan(text=input$txt2,quiet=TRUE,sep=",")
    myquant
  })
  
  #########
  output$txtresult <- renderText({
    mystock <-  mydata1()
    myquant <-  mydata2()
    
    code_quant = NULL
    for(i in 1:length(mystock)){
      if(mystock[i] %in% List_complete){
        code_quant[i] <- paste(c(mystock[i], as.character(cleaned[cleaned$code == mystock[i],colnames(cleaned)== "name"][1,]),
                                 myquant[i], "張" ,'\n'), collapse = " ")
      }else{
        code_quant <- as.character("抱歉，個股代碼輸入錯誤 或 該股資料筆數不足無法計算")
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
  
  
  #########  
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
      top10_hedge_weight[j] <- as.data.frame(hedge_weight[3,j] * cleaned[which(cleaned$"code"== top10_sum_corr[j]),6])  
    }
    top10_hedge_weight <- as.data.frame(top10_hedge_weight)
    colnames(top10_hedge_weight) <- c(top10_sum_corr)
    
    #避險後投組
    hedge_return = NULL
    for (j in 1:length(top10_sum_corr)) {
      for(h in seq(from = 1, to = 20, by = 2))
        hedge_return[j] <- top10_hedge_weight[j] + apply(mystock_hedgeWeight[h:(h+1)], 1, sum)
    }
    hedge_return <-  as.data.frame(hedge_return)
    colnames(hedge_return) <- c(top10_sum_corr)
    
    hedge_ret_sd <- data.frame(平均日報酬率 = apply(hedge_return, 2, mean),#投組平均日報酬(避險)
                                     平均日標準差 = apply(hedge_return, 2, sd))#投組平均日標準差(避險)
    hedge_portVaR <- data.frame(VaR = hedge_ret_sd$"平均日報酬率" - 1.65 * hedge_ret_sd$"平均日標準差") #投組VaR(避險)
    hedge_port <- data.frame(hedge_portVaR, hedge_ret_sd[1], 應購張數 = hedge_num)#避險結果
    
    #篩選避險後有改善風險之個股
    well_hedge <- hedge_port[which(hedge_port$VaR > portVaR),]
    if(nrow(well_hedge) == 0){
      cat("您的投資組合近乎完美!已找不到可加入之避險股!!",'\n',"再見");quit()
    }
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
    well_hedge <- datasetInput()
    
    names(stock) <- c("日期", "個股代碼", "簡稱", "產業別", "收盤價", "報酬率", "千成交量")
    df_well_hedge <- stock[which(stock$'個股代碼' %in% rownames(well_hedge) & stock$'日期' == max(stock$'日期')),c(2:5)]
    
    df_well_hedge
  })
  
  
  datasetInput3 <- reactive({ 
    well_hedge <- datasetInput()
    
    ret_sd <- data.frame((well_hedge$VaR) *-100,(well_hedge$"平均日報酬率")*100)
    colnames(ret_sd) <- c("日最大損失","日報酬率")
    rownames(ret_sd) <- row.names(well_hedge)
    ret_sd
  })
  
  
  #########  
  output$txtresult3 <- renderPrint({
    if (input$submitbutton>0) { 
      isolate("Calculation complete.") 
    } else {
      return("Server is ready for calculation.")
    }
  })
  
  output$plot <- renderPlot({
    if (input$submitbutton>0) { 
      isolate(datasetInput3())
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
  
  output$tabledata <- renderTable({
    if (input$submitbutton>0) { 
      isolate(datasetInput()) 
    } 
  },digits = 6)
  
  output$tabledata2 <- renderTable({
    if (input$submitbutton>0) { 
      isolate(datasetInput2()) 
    } 
  })
  
  output$img <- renderImage({
    list(src="industry.png", width="100%", alt = "錯誤")
  }, deleteFile=FALSE)
  
  output$img2 <- renderImage({
    list(src="dis_var.png", width="90%", alt = "錯誤")
  }, deleteFile=FALSE)
  
  output$img3 <- renderImage({
    list(src="formula.png", width="50%", alt = "錯誤")
  }, deleteFile=FALSE)
  
}


##shinyApp#########
shinyApp(ui, server)