################################################################################
#                                                                              #
#                           Student Finance Refund                             #    
#                                                                              #
################################################################################

# ----- Load Requirements ------------------------------------------------------

library(tidyverse)
library(readxl)
library(writexl)

library(shinydashboard)
library(DT)
library(shinythemes)

# ---------- Shiny Application -------------------------------------------------

# ---------- User Interface ----------------------------------------------------

# Dashboard
ui <- navbarPage(
  "National University of Singapore",
  theme=shinytheme("cosmo"),
  tabPanel(
    "Office of Finance",
    sidebarLayout(
      sidebarPanel(
        tags$head(
          tags$style(HTML("hr {border-top: 1px solid #000000;}"))
        ),
        p(strong("Refund List Generator")),
        hr(),
        fileInput("filelist", "File List (.csv)", multiple=FALSE, accept=".csv"),
        fileInput("blacklist1", "Blacklist 1 (.xlsx)", multiple=FALSE, accept=".xlsx"),
        fileInput("blacklist2", "Blacklist 2 (.xlsx)", multiple=FALSE, accept=".xlsx"),
        fileInput("refundlist", "Refund List (.xlsx)", multiple=FALSE, accept=".xlsx"),
        fileInput("feereport", "Fee Report (.xlsx)", multiple=FALSE, accept=".xlsx"),
        strong("Step 3: Download File with Class Membership"),
        br(),
        downloadButton("download", "Download"),
        hr(),
        p("Last Updated on 29 May 2022.")
      ),
      mainPanel(
        tabsetPanel(type="tabs",
                    tabPanel(
                      "Full Refund List", 
                      br(), br(),
                      dataTableOutput("fullRefundList",)
                    ),
                    tabPanel(
                      "Refund List with Remarks", 
                      br(), br(),
                      dataTableOutput("remarksRefundList",)
                    )
                    )
      )
    )
  )
)

# ---------- Server ------------------------------------------------------------

# Server
server <- function(input, output) {
  
  data <- reactive({
    
    if (is.null(input$filelist)) {
      return(NULL)
    }
    
    if (is.null(input$blacklist1)) {
      return(NULL)
    }
    
    if (is.null(input$blacklist2)) {
      return(NULL)
    }
    
    if (is.null(input$refundlist)) {
      return(NULL)
    }
    
    if (is.null(input$feereport)) {
      return(NULL)
    }
    
    inputFilelist <- input$filelist$datapath
    inputBlacklist1 <- input$blacklist1$datapath
    inputBlacklist2 <- input$blacklist2$datapath
    inputRefundlist <- input$refundlist$datapath
    inputFeereport <- input$feereport$datapath
    
    # Read File Names for Download
    file_list <- read.csv(inputFilelist)
    file_list$File.Name <- c("inputRefundlist", "inputFeereport", "inputBlacklist1", "inputBlacklist2")
    
    # Read the Blacklist
    category = "Blacklist"
    blacklist_filenames <- file_list %>%
      filter(Category==category) %>%
      select(File.Name, Remarks)
    refund_blacklist_dict <- list()
    for (i in blacklist_filenames$File.Name) {
      keyname <- blacklist_filenames %>%
        filter(File.Name==i) %>%
        select(Remarks) %>%
        toString()
      table <- read_xlsx(eval(parse(text=i)), skip=1)
      table <- table %>% mutate(Blacklist.Remarks=keyname)
      refund_blacklist_dict[[i]] <- table
    }
    blacklist <- bind_rows(refund_blacklist_dict)
    
    # Read the Refund List
    category = "Refund List"
    refund_list_filename <- file_list %>%
      filter(Category==category) %>%
      select(File.Name) %>%
      toString()
    refund_list <- read_xlsx(eval(parse(text=refund_list_filename)), skip=1)
    refund_list <- left_join(refund_list, blacklist, by=c("ID"="ID")) %>%
      arrange(desc(Amount))
    
    # Read the Fee Report
    category = "Fee Report"
    columns <- c("ID", "PreValue", "Tuition Fee", "Non MOE STF", "SSG Subsidies/TF Discount",
                 "IS16 & CPE1 GST", "MMFs", "MMFs Discount", "Hostel Fee", "Scholarship/Bursary",
                 "CPF/PSEA/MENDAKI", "FA Loan", "Payments  Y STUDENT", "Refund", "ACCOUNT BAL")
    fee_report_filename <- file_list %>%
      filter(Category==category) %>%
      select(File.Name) %>%
      toString()
    fee_report <- read_xlsx(eval(parse(text=fee_report_filename)), skip=1) %>%
      select(columns)
    fee_report$Check.Acct.Bal <- fee_report %>%
      select(c("PreValue", "Tuition Fee", "Non MOE STF", "SSG Subsidies/TF Discount", "IS16 & CPE1 GST",
               "MMFs", "MMFs Discount", "Hostel Fee", "Scholarship/Bursary","CPF/PSEA/MENDAKI",
               "FA Loan", "Payments  Y STUDENT", "Refund")) %>%
      apply(MARGIN=1, FUN=sum)
    fee_report$Difference <- fee_report$Check.Acct.Bal-fee_report$`ACCOUNT BAL`
    fee_report$Fee.Report.Remarks <- ifelse(fee_report$Difference>=0.01, "Account Balance does not tally. Account Balance supposed to be more.",
                                            ifelse(fee_report$Difference<=-0.01, "Account Balance does not tally. Account Balance supposed to be lesser.", NA))
    fee_report_error <- fee_report %>%
      filter(!is.na(Fee.Report.Remarks))
    
    # Output
    refund_list <- left_join(refund_list, fee_report_error[,c("ID", "Fee.Report.Remarks")], by=c("ID"="ID")) %>%
      arrange(desc(Amount))
    refund_list_dict <- list()
    refund_list_dict[["Full Refund List"]] <- refund_list
    refund_list_dict[["Refund List with Remarks"]] <- refund_list %>%
      filter(!is.na(Blacklist.Remarks) | !is.na(Fee.Report.Remarks))
    
    refund_list_dict
    
  })
  
  output$fullRefundList <- renderDataTable({
    data()[[1]]
  })
  
  output$remarksRefundList <- renderDataTable({
    data()[[2]]
  })

  output$download <- downloadHandler(
    filename = "Output_Refund List Summary.xlsx", 
    content = function(file) {
      write_xlsx(data(), path=file)
      }
  )
}

# ---------- Load Application --------------------------------------------------

shinyApp(ui, server)
