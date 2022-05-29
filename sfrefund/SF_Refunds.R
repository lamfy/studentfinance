# Load Libraries
library(tidyverse)
library(readxl)
library(openxlsx)

# Read File Names for Download
file_list <- read.csv("Download_file_names.csv")

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
  table <- read_xlsx(i, skip=1)
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
refund_list <- read_xlsx(refund_list_filename, skip=1)
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
fee_report <- read_xlsx(fee_report_filename, skip=1) %>%
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
write.xlsx(refund_list_dict, file="Output_Refund List Summary.xlsx")