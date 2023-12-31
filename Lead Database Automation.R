install.packages("openxlsx")
install.packages("writexl")
library(writexl)
library(bigQueryR)
library(bigrquery)
library(dplyr)
library(tidyr)
library(openxlsx)

bq_auth()

#Select the Table to use from Big Query
projectid = "your_project_id"

#Select Table 1 from Big Query
sql_1 <- "SELECT * FROM `your_project_id.client_name.table_name_1`"
tb_1 <- bq_project_query(projectid, sql_1)
table_1 <- bq_table_download(tb_1)

#Select Table 2 from Big Query
sql_2 <- "SELECT * FROM `your_project_id.client_name.table_name_2`"
tb_2 <- bq_project_query(projectid, sql_2)
table_2 <- bq_table_download(tb_2)

#Select Table 2 from Big Query
sql_3 <- "SELECT * FROM `your_project_id.client_name.table_name_3`"
tb_3 <- bq_project_query(projectid, sql_3)
table_3 <- bq_table_download(tb_3)

#Select Table 4 from Big Query
sql_4 <- "SELECT * FROM `your_project_id.client_name.table_name_4`"
tb_4 <- bq_project_query(projectid, sql_4)
table_4 <- bq_table_download(tb_4)


#Xlookups for old leads - Segment 1
old_leads <- inner_join(table_1, table_4, by=c('Case_Safe_Id'='_casesafeid'))

#xlookups for new leads - Segment 2
new_leads <- anti_join(table_1, table_4, by=c('Case_Safe_Id'='_casesafeid'))


#remove existing leads from Segment 1
old_leads_clean <- anti_join(old_leads, table_2, by=c('Case_Safe_Id'='Case_Safe_Id'))

#remove existing leads from Segment 2
new_leads_clean <- anti_join(new_leads, table_2, by=c('Case_Safe_Id'='Case_Safe_Id'))


#Identify Duplicate from Segment 1
duplicated_email_Segment_1 <- old_leads_clean[duplicated(old_leads_clean$Email)|
                                             duplicated(old_leads_clean$Email, 
                                                        fromLast=TRUE),]

#Identify Duplicate from Segment 2
duplicated_email_Segment_2 <- new_leads_clean[duplicated(new_leads_clean$Email)|
                                             duplicated(new_leads_clean$Email, 
                                                        fromLast=TRUE),]

#Remove duplicated in Segment 1
old_leads_clean_no_dupes <- old_leads_clean %>% anti_join(duplicated_email_Segment_1)

#Remove duplicated in Segment 2
new_leads_clean_no_dupes <- new_leads_clean %>% anti_join(duplicated_email_Segment_2)


#Merging Segment 1 to suppression list
old_leads_clean_no_dupes_suppressed <- 
  left_join(old_leads_clean_no_dupes,table_3,
            by=c('Case_Safe_Id' = '_key'))

#Merging Segment 2 to suppression list
new_leads_clean_no_dupes_suppressed <- 
  left_join(new_leads_clean_no_dupes,table_3,
            by=c('Case_Safe_Id' = '_key'))


#Remove Suppressed List in Segment 1
old_leads_clean_no_dupes_suppressed_rm <- old_leads_clean_no_dupes_suppressed %>% 
  filter(is.na(`_unsubscribereason`))

#Remove Suppressed List in Segment 2
new_leads_clean_no_dupes_suppressed_rm <- new_leads_clean_no_dupes_suppressed %>% 
  filter(is.na(`_unsubscribereason`))


#Extract Suppressed List
old_leads_clean_no_dupes_suppressedlist <- 
  inner_join(old_leads_clean_no_dupes,table_3,
             by=c('Case_Safe_Id' = '_key'))

#Extract Suppressed List
new_leads_clean_no_dupes_suppressedlist <- 
  inner_join(new_leads_clean_no_dupes,table_3,
             by=c('Case_Safe_Id' = '_key'))


#Select only column that required for Segment 1
Segment_1_Database <- old_leads_clean_no_dupes_suppressed_rm %>%
  select(`Full Name`=`name`,
         `...`,
         `...`, #Fill up your desired column here
         `...`)

#Select only column that required for Segment 2
Segment_2_Database <- new_leads_clean_no_dupes_suppressed_rm %>%
  select(`Full Name`=`name`,
         `...`,
         `...`, #Fill up your desired column here
         `...`)

#Select only column that required for Segment 1 Suppressed List
Segment_1_Suppressed <- old_leads_clean_no_dupes_suppressedlist %>%
  select(`Full Name`=`name`,
         `...`,
         `...`, #Fill up your desired column here
         `...`)

#Select only column that required for Segment 1 Duplicated
Segment_1_Duplicated <- duplicated_email_Segment_1 %>%
  select(`Full Name`=`name`,
         `...`,
         `...`, #Fill up your desired column here
         `...`)

#Select only column that required for Segment 2 Duplicated
Segment_2_Duplicated <- duplicated_email_Segment_2 %>%
  select(`Full Name`=`name`,
         `...`,
         `...`, #Fill up your desired column here
         `...`)

Summary <- data.frame(`Total SF Import` = nrow(table_1), 
                      "Segment 1" = nrow(Segment_1_Database),
                      "Segment 2" = nrow(Segment_2_Database),
                      "Suppressed List Segment 1" = nrow(Segment_1_Suppressed),
                      "Suppressed List Segment 2" = nrow(Segment_2_Suppressed),
                      "Duplicated List Segment 1" = nrow(Segment_1_Duplicated),
                      "Duplicated List Segment 2" = nrow(Segment_2_Duplicated))


Full_Database <- list(
  "Summary" = Summary,
  "Full Database" = table_1,
  "Segment 1" = Segment_1_Database,
  "Segment 2" = Segment_2_Database,
  "Suppressed List Segment 1" = Segment_1_Suppressed,
  "Suppressed List Segment 2" = Segment_2_Suppressed,
  "Duplicated List Segment 1" = Segment_1_Duplicated,
  "Duplicated List Segment 2" = Segment_2_Duplicated)

# Save the workbook to an Excel file
write.xlsx(Full_Database, file = "C:/Users/xxxxx/ xxxx/......./....... Automation/Full_Database.xlsx")

#Remove punctuation from Name - Segment 1
Segment_1_Database $`First Name` <- gsub("[[:punct:]]", "", Segment_1_Database $`First Name`)
Segment_1_Database $`Last Name` <- gsub("[[:punct:]]", "", Segment_1_Database $`Last Name`)
Segment_1_Database $`Territory` <- ifelse(is.na(Segment_1_Database$Territory), "No Territory", 
                                       Segment_1_Database$Territory)

#Remove punctuation from Name - Segment 2
Segment_2_Database $`First Name` <- gsub("[[:punct:]]", "", Segment_2_Database $`First Name`)
Segment_2_Database $`Last Name` <- gsub("[[:punct:]]", "", Segment_2_Database $`Last Name`)
Segment_2_Database $`Territory` <- ifelse(is.na(Segment_2_Database$Territory), "No Territory", 
                                       Segment_2_Database$Territory)

#Additional Column for SFMC
SFMC_add_column <- data.frame(
  Territory = c( `...`,
                 `...`, #Fill up your desired territory here
                 `...`),
  From_Name = c( `...`,
                 `...`, #Fill up your desired name here
                 `...`),
  From_Email = c( `...`,
                  `...`, #Fill up your desired email here
                  `...`)


#Xlookups for SFMC additional column - Segment 1
Segment_1_Database <- inner_join(Segment_1_Database, SFMC_add_column, 
                              by=c('Territory'='Territory'))

#Xlookups for SFMC additional column - Segment 2
Segment_2_Database <- inner_join(Segment_2_Database, SFMC_add_column, 
                              by=c('Territory'='Territory'))


#Extract column for SFMC File - Segment 1
Segment_1_SFMC<- Segment_1_Database %>%
  select("Case Safe Id" = `Case Safe ID`,
         `...`,
         `...`, #Fill up your desired column here
         `...`)
         "From Email" = `From_Email`)

#Extract column for SFMC File - Segment 2
Segment_2_SFMC<- Segment_2_Database %>%
  select("Case Safe Id" = `Case Safe ID`,
         `...`,
         `...`, #Fill up your desired column here
         `...`)
         "From Email" = `From_Email`)

#Specify the file path where you want to save the CSV file - Segment 1
file_path_1 <- "C:/Users/xxxxx/ xxxx/......./....... Automation/Segment_1_SFMC.csv"
write.csv(Segment_1_SFMC, file = file_path_1, row.names = FALSE)

#Specify the file path where you want to save the CSV file - Segment 2
file_path_2 <- "C:/Users/xxxxx/ xxxx/......./....... Automation/Segment_2_SFMC.csv"
write.csv(Segment_2_SFMC, file = file_path_2, row.names = FALSE)


#Form fill Combined Segment
Combined_Segment <- rbind(Segment_1_Database,Segment_2_Database)

#Select olumn to Use in Form Fill
Form_Fill_SFMC <- Combined_Segment %>%
  select("ID" = `Case Safe ID`,
         `...`,
         `...`, #Fill up your desired column here
         `...`)
         "Territory" = Territory)

#Create Blank Variables
Form_Fill_SFMC$"..." <- vector("character", length = nrow(Form_Fill_SFMC))
Form_Fill_SFMC$"..." <- vector("character", length = nrow(Form_Fill_SFMC))
Form_Fill_SFMC$"..." <- vector("character", length = nrow(Form_Fill_SFMC))

#Change the Naming Convention
Form_Fill_SFMC <- Form_Fill_SFMC %>%
  select("ID",
         `...`,
         `...`, #Fill up your desired column here
         `...`)
         "..." = `No_of_...`,
         "..." = `No_of_...`,
         "..." = `Plan...`)

#Specify the file path where you want to save the CSV file - Form Fill
file_path_3 <- "C:/Users/xxxxx/ xxxx/......./....... Automation/Form_Fill_SFMC.csv"
write.csv(Form_Fill_SFMC, file = file_path_3, row.names = FALSE)


