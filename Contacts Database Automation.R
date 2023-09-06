install.packages("openxlsx")
install.packages("writexl")
library(writexl)
library(bigrquery)
library(dplyr) #Include Dplyr Library for Lookups
library(openxlsx) # Create an Excel workbook with multiple sheets
library(tidyr)

bq_auth()

#Select the Project ID to use from Big Query
projectid = "your_project_id"

#Select Table 1 from Big Query
sql_1 <- "SELECT * FROM `your_project_id.client_name.table_name_1`"
tb_1 <- bq_project_query(projectid, sql_1)
table_1 <- bq_table_download(tb_1)

#Select Table 2 from Big Query
sql_2 <- "SELECT * FROM `your_project_id.client_name.table_name_2`"
tb_2 <- bq_project_query(projectid, sql_2)
table_2 <- bq_table_download(tb_2)

#Select Table 3 Table from Big Query
sql_3 <- "SELECT * FROM `your_project_id.client_name.table_name_3`"
tb_3 <- bq_project_query(projectid, sql_3)
table_3 <- bq_table_download(tb_3)

#Lookups For IDs from Table 2 to Table 1
table_1_new <- inner_join(table_1,table_2, 
                          by=c('contacts_id' = 'financial_id'))

#Identify duplicated in emails
duplicated_email <- table_1_new[duplicated(table_1_new$contacts_email)|
                                  duplicated(table_1_new$contacts_email, fromLast=TRUE),]

#Filter for rows with the same email but different IDs
same_email_diff_id <- duplicated_email %>%
  group_by(`contacts_email`) %>%
  filter(n_distinct(contacts_id) > 1)

# Filter for rows with the same email and the same ID
same_email_same_id <- duplicated_email %>%
  group_by(`contacts_email`) %>%
  filter(n_distinct(contacts_id) == 1)

#Remove duplicated in emails
table_1_new_cleaned <- table_1_new %>% anti_join(duplicated_email)

#Filter latest date from Table 3 
table_3_new <- table_3 %>% filter(`_date` == 09/06/20233)

#Lookups cleaned data from Table 1 to Table 3
table_1_new_cleaned_combined <- 
  left_join(table_1_new_cleaned,table_3_new,
            by=c('contacts_id' = 'leads_id'))
table_1_new_cleaned_combined <- table_1_new_cleaned_combined %>% distinct(contacts_id, .keep_all = TRUE)

#Extract Unique Leads from Table 1 that have in Table 3 (Based on Lookups with Table 3)
table_1_unique <- 
  inner_join(table_1_new_cleaned,table_3_new,
             by=c('contacts_id' = 'leads_id'))

#Remove Table 3 leads from Table 1
table_1_new_cleaned_combined <- table_1_new_cleaned_combined %>% filter(is.na(`table_3_unsubscribe`))

#Table 1 (Do Not Email)
table_1_Do_Not_Email  <- table_1_new_cleaned_combined %>% filter(`do_not_email` == TRUE)

#Remove Table 1 (Do Not Email)
table_1_new_cleaned_combined <- table_1_new_cleaned_combined %>% filter(`do_not_email` == FALSE)


#Separate into First Name and Last Name
table_1_new_cleaned_combined <- table_1_new_cleaned_combined %>%
  separate(`contacts_name`, into = c("First_Name", "Last_Name"), 
           sep = "\\s(?=[^\\s]+$)",
           remove = FALSE)

#Select only column that required for Contacts Database
Contacts_Database <- table_1_new_cleaned_combined %>%
  select(`Contacts ID`=`contacts_id`,
         `...`,
         `...`, #Fill up your desired column here
         `...`,
         `Salary` = `contacts_salary`)

Database_combined <- list(
  "Cleaned Database" = Contacts_Database,
  "Same Email Diff Id" = same_email_diff_id,
  "Same Email Same Id" = same_email_same_id,
  "Unique List" = table_1_unique)

# Save the workbook to an Excel file
write.xlsx(Database_combined, file = "C:/Users/xxxxx/ xxxx/......./....... Automation/Contacts_Database.xlsx")

#Remove punctuation from Name
Contacts_Database$`First Name` <- gsub("[[:punct:]]", "", Contacts_Database$`First Name`)
Contacts_Database$`Last Name` <- gsub("[[:punct:]]", "", Contacts_Database$`Last Name`)

#From the Contacts_Database, extract 4 column needed to SFMC csv file
Contacts_SFMC <- Contacts_Database %>%
  select(`Contacts ID`,
         `Last Name`,
         `First Name`,
         `Contacts Email`)

# Specify the file path where you want to save the CSV file
file_path <- "C:/Users/xxxxx/ xxxx/......./....... Automation/Contacts_SFMC.csv"

# Save the data frame to a CSV file
write.csv(Contacts_SFMC, file = file_path, row.names = FALSE)

# Display a message to confirm the file has been saved
cat("Data frame saved to CSV file successfully.\n")