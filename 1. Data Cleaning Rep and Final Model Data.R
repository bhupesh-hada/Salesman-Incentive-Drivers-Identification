
###### NEED TO UPDATE --- CHANGED COLUMN NAMES ######

###### NEED TO UPDATE --- THINK -- WILL WE EVER DO A MODEL REFRESH AND A MODEL RERUN AT THE SAME TIME? ######
###### NEED TO UPDATE --- THINK -- THIS IS DEPENDEDNT THAT AS WE PULL IN MORE DATA, THE PAST MONTHS DATA WONT CHANGE ELSE IT AFFECTS THE MODEL RERUN RESULTS ######
rm(list = ls())
library(dplyr)
library(data.table)
library(gtools)
library(xlsx)
library(bit64)
library(openxlsx)
library(tidyr)



###### << Manual>> Upload CSV ######
for(z in c(15)){
  print(z)#}
# For PH
`Con.Master.Data` <- read.csv("//falmumapp39/RTM Vietnam (19-AIM-3081)/SIO Analytics/PH/Data/03 Jan 2020/All Data/PH SIO Master Data 2020-01-06.csv", 
                               check.names = FALSE, encoding = "UTF-8")


###### << Manual>> Upload CSV for Mapping Needed Measures Names  ######

`Mapping.Needed.Measures.Names` <- read.csv("//falmumapp39/RTM Vietnam (19-AIM-3081)/SIO Analytics/Code/3. Modeling/Requirements Overall for Script - Mapping Needed Measures Names - 180726.csv", 
                                            check.names = FALSE, encoding = "UTF-8")




###### << Manual>> Select the Months ######
# Note PH Requirement to remove December is address in section "Standard (Add Column "Training / Testing / Remove Months)"
# Need to have a value for all the following (No NAs)

    # Start and End Month Needs to Emcompass all Months used 
    Start.Month = 2017.10
    End.Month = 2019.12
    
    # KPI Refresh Training Month and Testing Month Cannot Overlap 
    KPI.Refresh.Train.Start.Month = 2017.10
    KPI.Refresh.Train.End.Month = 2019.09
    
    KPI.Refresh.Test.Start.Month = 2019.10
    KPI.Refresh.Test.End.Month = 2019.12
    
    # Coefficient Refresh Training Month and Testing Month Cannot Overlap 
    # At Minimum, Coefficient Refresh full period needs to overlap with with KPI Refresh Full Period
    Coefficient.Refresh.Train.Start.Month = 2017.10
    Coefficient.Refresh.Train.End.Month = 2019.09
      
    Coefficient.Refresh.Test.Start.Month = 2019.10
    Coefficient.Refresh.Test.End.Month = 2019.12
    
    

###### Standard (Add Calculated Variables) * With Country Specific Coding ######
# Note PH and VN Requirement is for LE Online Report
# Note ID Requirement is for LE Offline Report and SAP 
    
if(unique(Con.Master.Data$`Country Code`) =="PH" | unique(Con.Master.Data$`Country Code`) =="VN"){
      Con.Master.Data$`LPPC` <- Con.Master.Data$`# SKUs(Order)`/Con.Master.Data$`# Productive Calls`
      Con.Master.Data$`% ECO to active shops` <- Con.Master.Data$`ECO Achieved`/Con.Master.Data$`No Of Month End Active Shop in PJP Covered`*100
      Con.Master.Data$`% BP to total calls` <- Con.Master.Data$`BP Achieved`/Con.Master.Data$`No Of Call  to be Visited`*100
      Con.Master.Data$`Average Daily BP` <- Con.Master.Data$`BP Achieved`/Con.Master.Data$`Total Working Days`
    }


    
######  Standard (Map Column Names) * With Country Specific Coding ######
# Note PH and VN Requirement is for LE Online Report
# Note ID Requirement is for LE Offline Report and SAP         
      
if(unique(Con.Master.Data$`Country Code`) =="PH" | unique(Con.Master.Data$`Country Code`) =="VN"){
      for(i in 1:nrow(Mapping.Needed.Measures.Names)){ 
        names(Con.Master.Data)[names(Con.Master.Data) == Mapping.Needed.Measures.Names$`Column Name LE Online`[i]] <- as.character(Mapping.Needed.Measures.Names$`Column Name SIO`[i])
      }
    }
  
remove(i)  
remove(Mapping.Needed.Measures.Names)
    


###### Standard (Add in "Portfolio Changes" for KPI and Coefficient Refresh) * With Country Specific Coding ######
# Note PH Requirement: Exclude distributor-site-salesperson which 
                        # Have a proportion outlet change of >20% for any month compared to the previous month 
                          ## excluding month of the start analysis period; or 
                          ## the first month the distributor-site-salesperson-sales.route.type appeared in the data
                        # Does not have data for at least 12?? (Play Around) consecutive months (Use minimum date not minimum months)
                          ## (therefore use na.rm = FALSE after putting that start analysis period month or first time the distributor-site-salesperson appeared in the data = 0)
                        # Is not present in the latest month of the analysis period 
# Note VN Requirement: Same Product Group for both training and testing period 
# Note ID Requirement: NIL          

Con.Master.Data.W.Portfolio <- Con.Master.Data



if(unique(Con.Master.Data$`Country Code`) =="PH"){
  
  Portfolio.Changes.KPI.Refresh <- subset(Con.Master.Data, Lag == 0 & (Month >= KPI.Refresh.Train.Start.Month & Month <= KPI.Refresh.Test.End.Month))
  Portfolio.Changes.Coefficient.Refresh <- subset(Con.Master.Data, Lag == 0 & (Month >= Coefficient.Refresh.Train.Start.Month & Month <= Coefficient.Refresh.Test.End.Month))
  
  Min.Number.of.Months.In.Sales.Route.Type <- z#8
  
  Portfolio.Changes.Type <- c("Portfolio.Changes.KPI.Refresh", "Portfolio.Changes.Coefficient.Refresh")
  
  ###### LOOP START > PH Portfolio Change ######
  
  for (Portfolio.Changes in Portfolio.Changes.Type) {
    # Portfolio.Changes = "Portfolio.Changes.KPI.Refresh"
    # Portfolio.Changes = "Portfolio.Changes.Coefficient.Refresh"
  
    Portfolio.Changes.Temp <- if (Portfolio.Changes == "Portfolio.Changes.KPI.Refresh") {Portfolio.Changes.KPI.Refresh} else
                                if (Portfolio.Changes == "Portfolio.Changes.Coefficient.Refresh") {Portfolio.Changes.Coefficient.Refresh} 
    
    Portfolio.Changes.Temp <- subset(Portfolio.Changes.Temp,
                                   select = c( "Country Code",      "Country", 
                                               "Distributor Code",  "Distributor",          "Site Code",     "Site",
                                               "Salesperson ID",    "Salesperson Name",     "Month",
                                               "Master Sales Route Type Code",              "Master Sales Route Type",
                                               "Proportion of Max Added or Removed Outlets vs Current Month" ))
  
    Portfolio.Changes.Temp <- unique(Portfolio.Changes.Temp)
  
    # Find the Min Month the distributor-site-salesperson-sales.route.type Appeared in the Data
    Portfolio.Changes.Temp <- merge(Portfolio.Changes.Temp, 
                                    Portfolio.Changes.Temp %>% 
                                      group_by ( `Country Code`,      `Country`, 
                                                 `Distributor Code`,  `Distributor`,          `Site Code`,     `Site`,
                                                 `Salesperson ID`,    `Salesperson Name`,
                                                 `Master Sales Route Type Code`,              `Master Sales Route Type`) %>% 
                                      summarise (`Min Month` = min(`Month`)),
                                    all = TRUE)
  
  # Change the Proportion of Changed Outlets to 0 for the starting analysis month and the Min Month the distributor-site-salesperson Appeared in the Data
  # So that it will be reflected as a numeric and will be considered and included when identify the maximum month change across the analysis period
  Portfolio.Changes.Temp$`Proportion of Max Added or Removed Outlets vs Current Month` <- as.character(Portfolio.Changes.Temp$`Proportion of Max Added or Removed Outlets vs Current Month`)
  Portfolio.Changes.Temp$`Proportion of Max Added or Removed Outlets vs Current Month` <- ifelse (Portfolio.Changes.Temp$Month == KPI.Refresh.Train.Start.Month |
                                                                                                    Portfolio.Changes.Temp$Month == Portfolio.Changes.Temp$`Min Month` , 
                                                                                                  0, Portfolio.Changes.Temp$`Proportion of Max Added or Removed Outlets vs Current Month`)
  
  # Identify Portfolio Changes
  Portfolio.Changes.Temp.Temp <- Portfolio.Changes.Temp %>% 
                                   group_by ( `Country Code`,      `Country`, 
                                             `Distributor Code`,  `Distributor`,          `Site Code`,     `Site`,
                                             `Salesperson ID`,    `Salesperson Name`,
                                             `Master Sales Route Type Code`,              `Master Sales Route Type`) %>% 
                                   summarise (`No. of Months in Sales Route Type` = length (unique(`Month`)),
                                              `Max Proportion Outlet Change across Months` = max(as.numeric(`Proportion of Max Added or Removed Outlets vs Current Month`), na.rm = FALSE),
                                              `Present in Latest Month` = ifelse (max(`Month`) == KPI.Refresh.Test.End.Month, "Present", "Not Present")) 
                                 
  Portfolio.Changes.Temp.Temp$`Portfolio Changes` = ifelse (Portfolio.Changes.Temp.Temp$`No. of Months in Sales Route Type` >= Min.Number.of.Months.In.Sales.Route.Type & 
                                                              Portfolio.Changes.Temp.Temp$`Max Proportion Outlet Change across Months` <= 0.20 &
                                                              Portfolio.Changes.Temp.Temp$`Present in Latest Month` == "Present",
                                                            "Minimal or No Change", "Do Not Include in Final Analysis")
                                                                              
  
  if (Portfolio.Changes == "Portfolio.Changes.KPI.Refresh") {names(Portfolio.Changes.Temp.Temp)[names(Portfolio.Changes.Temp.Temp) == "Portfolio Changes"]  <- "Portfolio Changes - KPI Refresh"} else 
    if (Portfolio.Changes == "Portfolio.Changes.Coefficient.Refresh") {names(Portfolio.Changes.Temp.Temp)[names(Portfolio.Changes.Temp.Temp) == "Portfolio Changes"]  <- "Portfolio Changes - Coefficient Refresh"}
  
  
  ###### LOOP > PH Portfolio Change  >> Create for Export >> For Data Cleansing Report >> Portfolio Changes   ######
  
  if (Portfolio.Changes == "Portfolio.Changes.KPI.Refresh") {Portfolio.Changes.KPI.Refresh <- Portfolio.Changes.Temp.Temp} else 
    if (Portfolio.Changes == "Portfolio.Changes.Coefficient.Refresh") {Portfolio.Changes.Coefficient.Refresh <- Portfolio.Changes.Temp.Temp}
  
  
  ###### LOOP > PH Portfolio Change  >> Map to Con Master Data   ######
  
  Con.Master.Data.W.Portfolio <- merge(Con.Master.Data.W.Portfolio,
                                       subset(Portfolio.Changes.Temp.Temp,
                                              select = -c( `No. of Months in Sales Route Type`,
                                                           `Max Proportion Outlet Change across Months`,
                                                           `Present in Latest Month`)),
                                       all = TRUE)
  
  rm(Portfolio.Changes.Temp)
  rm(Portfolio.Changes.Temp.Temp)
  
  }
  
  ###### LOOP END > PH Portfolio Change ######
  
  rm(Portfolio.Changes.Type)
  rm(Min.Number.of.Months.In.Sales.Route.Type)
  
  
  # Remove Columns in the Master Data that is used to derive "Portfolio Changes" to standardise the structure of reports in following sections
  ######  !!!! WIP !!!! Remove PJP Changes Here and columns  ###### 
    Con.Master.Data.W.Portfolio <-  subset(Con.Master.Data.W.Portfolio,
                                         select = -c( #`PJP Changes`,
                                                      `No  of Total Changed Outlets vs Current Month`,
                                                      `Proportion of Total Changed Outlets vs Current Month`,
                                                      `No  of Average Changed Outlets vs Current Month`,
                                                      `Proportion of Average Changed Outlets vs Current Month`,
                                                      `No  of Max Added or Removed Outlets vs Current Month`,
                                                      `Proportion of Max Added or Removed Outlets vs Current Month`))
  
 }
    
  


  
######  !!!! WIP !!!! Portfolio Changes VN Requirement: Same Product Group for both training and testing period ###### 
  # if(unique(Con.Master.Data$`Country Code`) =="VN"){ }


# Add in "Portfolio Changes"  Columns to standardise the structure of reports in following sections
if (is.null(Con.Master.Data.W.Portfolio$`Portfolio Changes - KPI Refresh`)) {Con.Master.Data.W.Portfolio$`Portfolio Changes - KPI Refresh` <- "Not Applicable for Country"} 
if (is.null(Con.Master.Data.W.Portfolio$`Portfolio Changes - Coefficient Refresh`)) {Con.Master.Data.W.Portfolio$`Portfolio Changes - Coefficient Refresh` <- "Not Applicable for Country"} 




######  Standard (Subset out only the dates required) ######

Con.All.Month <-  subset ( Con.Master.Data.W.Portfolio, Month >= Start.Month & Month<= End.Month )
rm(Start.Month)
rm(End.Month)

  

###### Standard: Remove Non Regular Distributor and Sites (for PH) ######
Con.All.Edited <-  subset ( Con.All.Month,  
                              `Site Code`!= "4B03" 
                            & `Site Code`!= "4007"
                            & `Site Code`!= "4065"
                            & `Site Code`!= "4A65"
                            & `Site Code`!= "4B65"
                            & `Site Code`!= "4078"
                            & `Site Code`!= "4079"
                            )

    
###### >> Create For Export >> For Data Cleansing Report >> View Extreme Outliers ######    
    

Report.Extreme.Outliers <- subset(Con.All.Edited, duplicate == 0 & Lag == 0 & (`% Target - GSV%` <= -1000 | `% Target - GSV%` >= 1000))
Report.Extreme.Outliers <- subset(Report.Extreme.Outliers,
                                          select = c( "Country Code",      "Country", 
                                                      "Distributor Code",  "Distributor",          "Site Code",     "Site",
                                                      "Salesperson ID",    "Salesperson Name",     "Month",
                                                      "Salesperson Clustering",
                                                      "GSV",          "Salesperson Target",  "% Target - GSV%"))
Report.Extreme.Outliers <- arrange (Report.Extreme.Outliers, `Month`, `Country`, `Distributor Code`, `Site Code`, `Salesperson ID`)

           

###### Standard (Make Add NA as a Factor in Cluster Column) ######
Con.All.Edited$`Salesperson Clustering` <- addNA(Con.All.Edited$`Salesperson Clustering`)


###### Standard (Remove Duplicates Salesman / Month) ######
Con.All.Edited <- subset(Con.All.Edited, duplicate == 0)


###### Standard (Remove Extreme Outliers: Only keep % Target GSV% between +/- 1000%) ######
Con.All.Edited$`% Target - GSV%` <- ifelse (Con.All.Edited$`% Target - GSV%` <= -1000 | Con.All.Edited$`% Target - GSV%` >= 1000,
                                           NA, Con.All.Edited$`% Target - GSV%`)




######## Standard (Create Lag Columns) ########

rm(Con.All.Month)

# Subset Time Lag and Change Column Name
Full.Data.TimeLag0 <- Con.All.Edited[Con.All.Edited$Lag == 0, ]
Full.Data.TimeLag1 <- Con.All.Edited[Con.All.Edited$Lag == 1, ]
names (`Full.Data.TimeLag1`) <- paste0(names (`Con.All.Edited`), " (Lag -1 Month)")
Full.Data.TimeLag2 <- Con.All.Edited[Con.All.Edited$Lag == 2, ]
names (`Full.Data.TimeLag2`) <- paste0(names (`Con.All.Edited`), " (Lag -2 Month)")
Full.Data.TimeLag3 <- Con.All.Edited[Con.All.Edited$Lag == 3, ]
names (`Full.Data.TimeLag3`) <- paste0(names (`Con.All.Edited`), " (Lag -3 Month)")

# Merge the tables to add Time Lag Columns
Full.Data.TimeLag.All <-  merge ( merge ( merge ( Full.Data.TimeLag0, Full.Data.TimeLag1, all.x = TRUE,
                                                  by.x = c ("Country Code",                                             "Country", 
                                                            "Distributor Code",                                         "Distributor",    
                                                            "Site Code",                                                "Site",
                                                            "Salesperson ID",                                           "Salesperson Name",      
                                                            "Month",                                                    "Portfolio Changes - KPI Refresh",                
                                                            "Portfolio Changes - Coefficient Refresh",                  "Included in Analysis",   
                                                            "Salesperson Clustering",                                   "duplicate"),
                                                  by.y = c ("Country Code (Lag -1 Month)",                              "Country (Lag -1 Month)", 
                                                            "Distributor Code (Lag -1 Month)",                          "Distributor (Lag -1 Month)",    
                                                            "Site Code (Lag -1 Month)",                                 "Site (Lag -1 Month)",
                                                            "Salesperson ID (Lag -1 Month)",                            "Salesperson Name (Lag -1 Month)",      
                                                            "Month (Lag -1 Month)",                                     "Portfolio Changes - KPI Refresh (Lag -1 Month)", 
                                                            "Portfolio Changes - Coefficient Refresh (Lag -1 Month)",   "Included in Analysis (Lag -1 Month)",  
                                                            "Salesperson Clustering (Lag -1 Month)",                    "duplicate (Lag -1 Month)")),
                                          Full.Data.TimeLag2, all.x = TRUE,
                                          by.x = c ("Country Code",                                             "Country", 
                                                    "Distributor Code",                                         "Distributor",    
                                                    "Site Code",                                                "Site",
                                                    "Salesperson ID",                                           "Salesperson Name",      
                                                    "Month",                                                    "Portfolio Changes - KPI Refresh",                
                                                    "Portfolio Changes - Coefficient Refresh",                  "Included in Analysis",   
                                                    "Salesperson Clustering",                                   "duplicate"),
                                          by.y = c ("Country Code (Lag -2 Month)",                              "Country (Lag -2 Month)", 
                                                    "Distributor Code (Lag -2 Month)",                          "Distributor (Lag -2 Month)",    
                                                    "Site Code (Lag -2 Month)",                                 "Site (Lag -2 Month)",
                                                    "Salesperson ID (Lag -2 Month)",                            "Salesperson Name (Lag -2 Month)",      
                                                    "Month (Lag -2 Month)",                                     "Portfolio Changes - KPI Refresh (Lag -2 Month)", 
                                                    "Portfolio Changes - Coefficient Refresh (Lag -2 Month)",   "Included in Analysis (Lag -2 Month)",  
                                                    "Salesperson Clustering (Lag -2 Month)",                    "duplicate (Lag -2 Month)")),
                                  Full.Data.TimeLag3, all.x = TRUE,
                                  by.x = c ("Country Code",                                             "Country", 
                                            "Distributor Code",                                         "Distributor",    
                                            "Site Code",                                                "Site",
                                            "Salesperson ID",                                           "Salesperson Name",      
                                            "Month",                                                    "Portfolio Changes - KPI Refresh",                
                                            "Portfolio Changes - Coefficient Refresh",                  "Included in Analysis",   
                                            "Salesperson Clustering",                                   "duplicate"),
                                  by.y = c ("Country Code (Lag -3 Month)",                              "Country (Lag -3 Month)", 
                                            "Distributor Code (Lag -3 Month)",                          "Distributor (Lag -3 Month)",    
                                            "Site Code (Lag -3 Month)",                                 "Site (Lag -3 Month)",
                                            "Salesperson ID (Lag -3 Month)",                            "Salesperson Name (Lag -3 Month)",      
                                            "Month (Lag -3 Month)",                                     "Portfolio Changes - KPI Refresh (Lag -3 Month)", 
                                            "Portfolio Changes - Coefficient Refresh (Lag -3 Month)",   "Included in Analysis (Lag -3 Month)",  
                                            "Salesperson Clustering (Lag -3 Month)",                    "duplicate (Lag -3 Month)"))

Full.Data.TimeLag.All <- subset (Full.Data.TimeLag.All, 
                                 select = -c(`Lag`, `Lag (Lag -1 Month)`, `Lag (Lag -2 Month)`, `Lag (Lag -3 Month)`, duplicate))

rm(Full.Data.TimeLag0)
rm(Full.Data.TimeLag1)
rm(Full.Data.TimeLag2)
rm(Full.Data.TimeLag3)



###### Standard (Add Columns to identify Training / Testing / Removed Months) * w PH specific coding ######
# One Column for KPI Refresh Training / Testing / Removed Months 
# One Column for Coefficient Refresh Training / Testing / Removed Months
# Note PH Specific Code for Removed Months

Full.Data.All.Columns <- Full.Data.TimeLag.All

# One Column for KPI Refresh Training / Testing / Removed Months 
    Full.Data.All.Columns$`Month Grouping - KPI Refresh` <- ifelse(Full.Data.All.Columns$Month >= KPI.Refresh.Train.Start.Month &
                                                                     Full.Data.All.Columns$Month <= KPI.Refresh.Train.End.Month,
                                                                   "KPI Refresh - Training Months", 
                                                                   ifelse (Full.Data.All.Columns$Month >= KPI.Refresh.Test.Start.Month & 
                                                                             Full.Data.All.Columns$Month <= KPI.Refresh.Test.End.Month,
                                                                           "KPI Refresh - Testing Months",
                                                                           "KPI Refresh - Removed Months"))
    # ** PH Specific Requirement
    Full.Data.All.Columns$`Month Grouping - KPI Refresh`<- ifelse(Full.Data.All.Columns$`Country Code`=="PH" & 
                                                                    substring(Full.Data.All.Columns$Month,6,7) == 12,
                                                                  "KPI Refresh - Removed Months",
                                                                  Full.Data.All.Columns$`Month Grouping - KPI Refresh`)
    

# One Column for Coefficient Refresh Training / Testing / Removed Months
    Full.Data.All.Columns$`Month Grouping - Coefficient Refresh` <- ifelse(Full.Data.All.Columns$Month >= Coefficient.Refresh.Train.Start.Month &
                                                                     Full.Data.All.Columns$Month <= Coefficient.Refresh.Train.End.Month,
                                                                   "Coefficient Refresh - Training Months", 
                                                                   ifelse (Full.Data.All.Columns$Month >= Coefficient.Refresh.Test.Start.Month & 
                                                                             Full.Data.All.Columns$Month <= Coefficient.Refresh.Test.End.Month,
                                                                           "Coefficient Refresh - Testing Months",
                                                                           "Coefficient Refresh - Removed Months"))
    # ** PH Specific Requirement
    Full.Data.All.Columns$`Month Grouping - Coefficient Refresh`<- ifelse(Full.Data.All.Columns$`Country Code`=="PH" & 
                                                                    substring(Full.Data.All.Columns$Month,6,7) == 12,
                                                                  "Coefficient Refresh - Removed Months",
                                                                  Full.Data.All.Columns$`Month Grouping - Coefficient Refresh`)

rm(Full.Data.TimeLag.All)  



###### Standard (Subset Out Columns / KPIs / Measures Needed ) ######
# Note: KPIs / Measures pulled out from data dictionary

Full.Data.Needed.Columns <- subset (Full.Data.All.Columns, 
                                    select = c( "Country Code",                         "Country", 
                                                "Distributor Code",                     "Distributor",    
                                                "Site Code",                            "Site",
                                                "Salesperson ID",                       "Salesperson Name",      
                                                "Month",                                
                                                "Month Grouping - KPI Refresh",         "Month Grouping - Coefficient Refresh", 
                                                "Portfolio Changes - KPI Refresh",      "Portfolio Changes - Coefficient Refresh",                 
                                                "Included in Analysis",                 "Salesperson Clustering",
                                                "Appear in all reports",          
                                                "GSV",                                  "Salesperson Target",
                                                "% Target - GSV%",                   
                                                "Total Working Days",                   "Total Time Spent (Mins)",  
                                                "Effective Outlet Time (Mins)",         "Transit Time (Mins)",
                                                "# Planned Calls",                      "# Actual Calls",
                                                "# PJP Complied Calls",                 "% PJP Complied Calls",
                                                "# PJP Complied Days",                  "# Geo Complied Calls",
                                                "% Geo Complied Calls",                 "# Geo Complied Days",
                                                "No Of Call  to be Visited",            "No Of Month End Active Shop in PJP Covered",
                                                "% ECO to Active Outlet",               "# ECO Achieved",
                                                "% BP to Total Calls",                  "# BP Achieved",
                                                "# PJP Complied Productive Calls",      "Average Daily BP",
                                                "# SKUs(Order)",                        "# SKU's /Actual Calls",
                                                "CotC Lines",                           "CotC Actual",
                                                "CotC Achieved %",                      "Total Outlet Count",
                                                "# 4P Compliance Outlet",               "% Perfect Outlet Count",
                                                "EB Lines",                             "EB Actual",
                                                "EB Achieved %",                        "EB - Perfect Outlet Count",
                                                "Red Lines",                            "RL Actual",
                                                "RL Achieved %",                        "RL - Perfect Outlet Count",
                                                "NPD Lines",                            "NPD Actual",
                                                "NPD Achieved %",                       "NPD - Perfect Outlet Count",
                                                "WP Lines",                             "WP Actual",
                                                "WP Achieved %",                        "WP - Perfect Outlet Count"))



###### LOOP START > Create Loop for Clean Data For Modelling and Data Cleansing Biz Rep  ######

Model.Type <- c("Model.Rerun.KPI", "Model.Refresh.Coefficient")

for (Model.Type.Data in Model.Type) {
# Model.Type.Data = "Model.Rerun.KPI"
# Model.Type.Data = "Model.Refresh.Coefficient"

  
# This creates the columns "`Training / Testing / Removed Months" and "Portfolio Changes" use the correct column depending on what data being creating

Full.Data.Needed.Columns$`Training / Testing / Removed Months` <- if (Model.Type.Data == "Model.Rerun.KPI") {Full.Data.Needed.Columns$`Month Grouping - KPI Refresh`} else
  if (Model.Type.Data == "Model.Refresh.Coefficient") {Full.Data.Needed.Columns$`Month Grouping - Coefficient Refresh`} 
  
Full.Data.Needed.Columns$`Portfolio Changes` <- if (Model.Type.Data == "Model.Rerun.KPI") {Full.Data.Needed.Columns$`Portfolio Changes - KPI Refresh`} else
  if (Model.Type.Data == "Model.Refresh.Coefficient") {Full.Data.Needed.Columns$`Portfolio Changes - Coefficient Refresh`} 




###### LOOP > Full Data ######
Full.Data <- Full.Data.Needed.Columns



###### LOOP > Full Data (GSV) ######
Full.Data.GSV <-  subset ( Full.Data,  !is.na(GSV))
Full.Data.GSV <-  subset ( Full.Data.GSV, !is.na(`% Target - GSV%`))



###### LOOP > Full Data (In All Reports) ######
Full.Data.In.All.Reports <- subset (Full.Data, `Appear in all reports`== "In all reports")



###### LOOP > Full Data (In All Reports) ######

###### LOOP > In Analysis w Portfolio Change ######
In.Analysis.Data.Portfolio.Change <- Full.Data.In.All.Reports
In.Analysis.Data.Portfolio.Change <- subset (In.Analysis.Data.Portfolio.Change, `Portfolio Changes` == "Minimal or No Change" | `Portfolio Changes` == "Not Applicable for Country")



###### LOOP > Interim for Uncleaned 6sd: Subset Out Training + Testing Months ######
In.Analysis.Train.and.Test.Months <- rbind ( subset (In.Analysis.Data.Portfolio.Change, `Training / Testing / Removed Months` == "KPI Refresh - Training Months"),
                                             subset (In.Analysis.Data.Portfolio.Change, `Training / Testing / Removed Months` == "KPI Refresh - Testing Months"),
                                             subset (In.Analysis.Data.Portfolio.Change, `Training / Testing / Removed Months` == "Coefficient Refresh - Training Months"),
                                             subset (In.Analysis.Data.Portfolio.Change, `Training / Testing / Removed Months` == "Coefficient Refresh - Testing Months"))



###### LOOP > Interim for Uncleaned 6sd: In Analysis uncleaned within 6sd for Training + Testing Months ######
# Note Out of Analysis Clusters are removed from here onwards
# Note Out of Analysis Months are removed from here onwards

# Get 6sd for training + testing months 
In.Analysis.Data.Train.and.Test.Months.Uncleaned.6sd <- NULL

# Get Clusters for Loop

Salesperson.Clustering <- subset (In.Analysis.Train.and.Test.Months, `Included in Analysis` == "Y")
Salesperson.Clustering <- as.character(na.omit(as.character(unique(Salesperson.Clustering$`Salesperson Clustering`))))


# Run Loop to get 6sd
system.time ( for (Selected.Cluster in Salesperson.Clustering) { 
  
  print (Selected.Cluster)
    For.Loop.Cluster <- subset ( In.Analysis.Train.and.Test.Months,`Salesperson Clustering` %in% Selected.Cluster)
  
  # Look only at Measures
  For.Loop.Cluster.Measure <- subset (For.Loop.Cluster, 
                                      select = c(  "GSV",                                  "Salesperson Target",
                                                   "% Target - GSV%",                   
                                                   "Total Working Days",                   "Total Time Spent (Mins)",  
                                                   "Effective Outlet Time (Mins)",         "Transit Time (Mins)",
                                                   "# Planned Calls",                      "# Actual Calls",
                                                   "# PJP Complied Calls",                 "% PJP Complied Calls",
                                                   "# PJP Complied Days",                  "# Geo Complied Calls",
                                                   "% Geo Complied Calls",                 "# Geo Complied Days",
                                                   "No Of Call  to be Visited",            "No Of Month End Active Shop in PJP Covered",
                                                   "% ECO to Active Outlet",               "# ECO Achieved",
                                                   "% BP to Total Calls",                  "# BP Achieved",
                                                   "# PJP Complied Productive Calls",      "Average Daily BP",
                                                   "# SKUs(Order)",                        "# SKU's /Actual Calls",
                                                   "CotC Lines",                           "CotC Actual",
                                                   "CotC Achieved %",                      "Total Outlet Count",
                                                   "# 4P Compliance Outlet",               "% Perfect Outlet Count",
                                                   "EB Lines",                             "EB Actual",
                                                   "EB Achieved %",                        "EB - Perfect Outlet Count",
                                                   "Red Lines",                            "RL Actual",
                                                   "RL Achieved %",                        "RL - Perfect Outlet Count",
                                                   "NPD Lines",                            "NPD Actual",
                                                   "NPD Achieved %",                       "NPD - Perfect Outlet Count",
                                                   "WP Lines",                             "WP Actual",
                                                   "WP Achieved %",                        "WP - Perfect Outlet Count"))
  
  # Remove Outliers from Measures 
  # For apply (dataframe,2,function(x) "2" indicates look at columns
  # Calculate mean and sd for each column
  # For each column change x to na if na or outside 6sd
  For.Loop.Cluster.Measure.6sd <- apply (For.Loop.Cluster.Measure,2,function(x){ 
    colm  <-mean(x, na.rm = T)
    std   <-sd(x,na.rm = T)
    sapply(x,function(x) {
      if(x > colm+3*std | x < colm - 3*std | is.na(x)){
        x<-NA 
      }else{
        return(x)
      }
    })
  })
  
  # Combine Final Measures + GSV with Descriptive Variables
  For.Loop.Cluster.6sd <- data.frame ( cbind ( subset (For.Loop.Cluster, 
                                                       select = c( "Country Code",                         "Country", 
                                                                   "Distributor Code",                     "Distributor",    
                                                                   "Site Code",                            "Site",
                                                                   "Salesperson ID",                       "Salesperson Name",      
                                                                   "Month",                                "Month Grouping - KPI Refresh",       
                                                                   "Month Grouping - Coefficient Refresh", "Training / Testing / Removed Months",
                                                                   "Portfolio Changes - KPI Refresh",      "Portfolio Changes - Coefficient Refresh",
                                                                   "Portfolio Changes",                    "Included in Analysis",                 
                                                                   "Salesperson Clustering",               "Appear in all reports")),
                                               For.Loop.Cluster.Measure.6sd), check.names = FALSE)
  
  # Rbind the cluster data to the final data set
  In.Analysis.Data.Train.and.Test.Months.Uncleaned.6sd <- rbind(In.Analysis.Data.Train.and.Test.Months.Uncleaned.6sd,For.Loop.Cluster.6sd)
  
  rm(Selected.Cluster)
  rm(For.Loop.Cluster)
  rm(For.Loop.Cluster.Measure)
  rm(For.Loop.Cluster.Measure.6sd)
  rm(For.Loop.Cluster.6sd)
  
})

rm(Salesperson.Clustering)
rm(In.Analysis.Train.and.Test.Months)



###### LOOP > In Analysis w Portfolio Change within 6sd but Uncleaned ######
In.Analysis.Data.Portfolio.Change.6sd.Uncleaned <-  In.Analysis.Data.Train.and.Test.Months.Uncleaned.6sd
rm(In.Analysis.Data.Train.and.Test.Months.Uncleaned.6sd)



###### LOOP > In Analysis w Portfolio Change within 6sd Cleaned w Biz Logic ######
In.Analysis.Data.Portfolio.Change.6sd.Biz.Cleaned <- In.Analysis.Data.Portfolio.Change.6sd.Uncleaned

# Effective Outlet Time in (Mins): Effective Outlet Time in (Mins) > 25 mins
In.Analysis.Data.Portfolio.Change.6sd.Biz.Cleaned$`Effective Outlet Time (Mins)` <- ifelse (In.Analysis.Data.Portfolio.Change.6sd.Biz.Cleaned$`Effective Outlet Time (Mins)` < 25,
                                                                                            NA, In.Analysis.Data.Portfolio.Change.6sd.Biz.Cleaned$`Effective Outlet Time (Mins)`)          



###### LOOP > In Analysis w Portfolio Change within 6sd Cleaned w Biz Logic & Generic Logic ######
# (See Data Dictionary. Note Sequence is very important)
In.Analysis.Data.6sd.Biz.Generic.Cleaned <- In.Analysis.Data.Portfolio.Change.6sd.Biz.Cleaned



# Salesman Target: Salesman Target = NA when a) Salesman Target = 0
In.Analysis.Data.6sd.Biz.Generic.Cleaned$`Salesperson Target`<- ifelse ( In.Analysis.Data.6sd.Biz.Generic.Cleaned$`Salesperson Target` == 0 ,
                                                                      NA, In.Analysis.Data.6sd.Biz.Generic.Cleaned$`Salesperson Target`)
# % Target - GSV%: % Target - GSV = NA when a) Salesman Target = NA | b) GSV = NA 
In.Analysis.Data.6sd.Biz.Generic.Cleaned$`% Target - GSV%` <- ifelse ((is.na(In.Analysis.Data.6sd.Biz.Generic.Cleaned$`Salesperson Target`)| 
                                                                         is.na(In.Analysis.Data.6sd.Biz.Generic.Cleaned$`GSV`)),
                                                                      NA, In.Analysis.Data.6sd.Biz.Generic.Cleaned$`% Target - GSV%`)



# Effective Outlet Time in (Mins): Effective Outlet Time in (Mins) = NA when a) Effective Outlet Time in (Mins) <= 0
In.Analysis.Data.6sd.Biz.Generic.Cleaned$`Effective Outlet Time (Mins)`[In.Analysis.Data.6sd.Biz.Generic.Cleaned$`Effective Outlet Time (Mins))`<= 0] <- NA
# Transit Time (In Mins) : Transit Time (In Mins)  = NA when a) Transit Time (In Mins) <= 0
In.Analysis.Data.6sd.Biz.Generic.Cleaned$`Transit Time (Mins)`[In.Analysis.Data.6sd.Biz.Generic.Cleaned$`Transit Time (Mins)`<= 0] <- NA
# Total Time Spent  (in Mins): Total Time Spent  (in Mins) = NA when a) Effective Outlet Time in (Mins) = NA | b) Total Time Spent  (in Mins) = NA
In.Analysis.Data.6sd.Biz.Generic.Cleaned$`Total Time Spent (Mins)` <- ifelse (is.na (In.Analysis.Data.6sd.Biz.Generic.Cleaned$`Effective Outlet Time (Mins)`)|
                                                                                    is.na(In.Analysis.Data.6sd.Biz.Generic.Cleaned$`Transit Time (Mins)`),
                                                                                  NA, In.Analysis.Data.6sd.Biz.Generic.Cleaned$`Total Time Spent (Mins)`)



# PJP Complied Call: # PJP Complied Call = NA when a) # Planned Calls = NA 
In.Analysis.Data.6sd.Biz.Generic.Cleaned$`# PJP Complied Calls` <- ifelse ( is.na (In.Analysis.Data.6sd.Biz.Generic.Cleaned$`# Planned Calls`),
                                                                           NA, In.Analysis.Data.6sd.Biz.Generic.Cleaned$`# PJP Complied Calls`)
# % PJP Complied Calls: % PJP Complied Calls = NA when a) Planned Calls = NA | b) # PJP Complied Calls = NA
In.Analysis.Data.6sd.Biz.Generic.Cleaned$`% PJP Complied Calls` <- ifelse (is.na (In.Analysis.Data.6sd.Biz.Generic.Cleaned$`# Planned Calls`)|
                                                                             is.na(In.Analysis.Data.6sd.Biz.Generic.Cleaned$`# PJP Complied Calls`),
                                                                           NA, In.Analysis.Data.6sd.Biz.Generic.Cleaned$`% PJP Complied Calls`)
# # PJP Complied Days: # PJP Complied Days = NA when a) Planned Calls = NA | b) # PJP Complied Calls = NA
In.Analysis.Data.6sd.Biz.Generic.Cleaned$`# PJP Complied Days` <- ifelse (is.na (In.Analysis.Data.6sd.Biz.Generic.Cleaned$`# Planned Calls`)|
                                                                            is.na(In.Analysis.Data.6sd.Biz.Generic.Cleaned$`# PJP Complied Calls`),
                                                                          NA, In.Analysis.Data.6sd.Biz.Generic.Cleaned$`# PJP Complied Days`)



# Geo Complied Calls: # Geo Complied Calls = NA when a) # Actual Calls = NA 
In.Analysis.Data.6sd.Biz.Generic.Cleaned$`# Geo Complied Calls` <- ifelse ( is.na (In.Analysis.Data.6sd.Biz.Generic.Cleaned$`# Actual Calls`),
                                                                            NA, In.Analysis.Data.6sd.Biz.Generic.Cleaned$`# Geo Complied Calls`)
# % Geo Complied Calls: % Geo Complied Calls = NA when a) # Actual Calls = NA = NA | b) # Geo Complied Calls = NA
In.Analysis.Data.6sd.Biz.Generic.Cleaned$`% Geo Complied Calls` <- ifelse (is.na (In.Analysis.Data.6sd.Biz.Generic.Cleaned$`# Actual Calls`)|
                                                                             is.na(In.Analysis.Data.6sd.Biz.Generic.Cleaned$`# Geo Complied Calls`),
                                                                           NA, In.Analysis.Data.6sd.Biz.Generic.Cleaned$`% Geo Complied Calls` )
# # Geo Complied Days: # Geo Complied Days = NA when a) # Actual Calls = NA = NA | b) # Geo Complied Calls = NA
In.Analysis.Data.6sd.Biz.Generic.Cleaned$`# Geo Complied Days` <- ifelse (is.na (In.Analysis.Data.6sd.Biz.Generic.Cleaned$`# Actual Calls`)|
                                                                            is.na(In.Analysis.Data.6sd.Biz.Generic.Cleaned$`# Geo Complied Calls`),
                                                                          NA, In.Analysis.Data.6sd.Biz.Generic.Cleaned$`# Geo Complied Days`)



# ECO Achieved: ECO Achieved = NA when a) No Of Month End Active Shop in PJP Covered = NA | b) ECO Achieved < 0 
In.Analysis.Data.6sd.Biz.Generic.Cleaned$`# ECO Achieved` <- ifelse (is.na (In.Analysis.Data.6sd.Biz.Generic.Cleaned$`No Of Month End Active Shop in PJP Covered`)|
                                                                     In.Analysis.Data.6sd.Biz.Generic.Cleaned$`# ECO Achieved` < 0,
                                                                      NA, In.Analysis.Data.6sd.Biz.Generic.Cleaned$`# ECO Achieved`)
# % ECO to active shops (Cal.): % ECO to active shops (Cal.) = NA when a) No Of Month End Active Shop in PJP Covered = NA | b) ECO Achieved = NA
In.Analysis.Data.6sd.Biz.Generic.Cleaned$`% ECO to Active Outlet` <- ifelse (is.na (In.Analysis.Data.6sd.Biz.Generic.Cleaned$`No Of Month End Active Shop in PJP Covered`)|
                                                                                     is.na(In.Analysis.Data.6sd.Biz.Generic.Cleaned$`# ECO Achieved`),
                                                                                   NA, In.Analysis.Data.6sd.Biz.Generic.Cleaned$`% ECO to Active Outlet`)



# # BP Achieved : # BP Achieved  = NA when a) No Of Call  to be Visited  = NA
In.Analysis.Data.6sd.Biz.Generic.Cleaned$`# BP Achieved` <- ifelse (is.na (In.Analysis.Data.6sd.Biz.Generic.Cleaned$`No Of Call  to be Visited`),
                                                                  NA, In.Analysis.Data.6sd.Biz.Generic.Cleaned$`# BP Achieved`)
# % BP to total calls (Cal.): % BP to total calls  = NA when a) No Of Call  to be Visited = NA | b) # BP Achieved = NA
In.Analysis.Data.6sd.Biz.Generic.Cleaned$`% BP to Total Calls` <- ifelse (is.na (In.Analysis.Data.6sd.Biz.Generic.Cleaned$`No Of Call  to be Visited`)|
                                                                                   is.na(In.Analysis.Data.6sd.Biz.Generic.Cleaned$`# BP Achieved`),
                                                                                 NA, In.Analysis.Data.6sd.Biz.Generic.Cleaned$`% BP to Total Calls`)
# Average Daily BP (Cal.): Average Daily BP  = NA when a) Total Working Days = NA | b) # BP Achieved = NA
In.Analysis.Data.6sd.Biz.Generic.Cleaned$`Average Daily BP` <- ifelse (is.na (In.Analysis.Data.6sd.Biz.Generic.Cleaned$`Total Working Days`)|
                                                                                is.na(In.Analysis.Data.6sd.Biz.Generic.Cleaned$`# BP Achieved`),
                                                                              NA, In.Analysis.Data.6sd.Biz.Generic.Cleaned$`Average Daily BP`)



# # PJP Complied Productive Calls : # PJP Complied Productive Calls = NA when a) # PJP Complied Call = NA
In.Analysis.Data.6sd.Biz.Generic.Cleaned$`# PJP Complied Productive Calls` <- ifelse (is.na (In.Analysis.Data.6sd.Biz.Generic.Cleaned$`# PJP Complied Call`),
                                                                                      NA, In.Analysis.Data.6sd.Biz.Generic.Cleaned$`# PJP Complied Productive Calls`)



# # SKU's /Actual Calls: # SKU's /Actual Calls = NA when a)  # SKUs(Order) = NA | b) # Actual Calls  = NA  
In.Analysis.Data.6sd.Biz.Generic.Cleaned$`# SKU's /Actual Calls` <- ifelse (is.na (In.Analysis.Data.6sd.Biz.Generic.Cleaned$`# SKUs(Order)`)|
                                                                              is.na(In.Analysis.Data.6sd.Biz.Generic.Cleaned$`# Actual Calls`),
                                                                            NA, In.Analysis.Data.6sd.Biz.Generic.Cleaned$`# SKU's /Actual Calls`)



# CotC Actual:  CotC Actual = NA when a) CotC Lines = NA | b) CotC Target <  CotC Actual
In.Analysis.Data.6sd.Biz.Generic.Cleaned$`CotC Actual` <- ifelse (is.na (In.Analysis.Data.6sd.Biz.Generic.Cleaned$`CotC Lines`)|
                                                                      In.Analysis.Data.6sd.Biz.Generic.Cleaned$`CotC Lines`< In.Analysis.Data.6sd.Biz.Generic.Cleaned$`CotC Actual`,
                                                                    NA, In.Analysis.Data.6sd.Biz.Generic.Cleaned$`CotC Actual`)
# CotC Achieved %: CotC Achieved % = NA when a) CotC Lines = NA | b) CotC Actual = NA
In.Analysis.Data.6sd.Biz.Generic.Cleaned$`CotC Achieved %` <- ifelse (is.na (In.Analysis.Data.6sd.Biz.Generic.Cleaned$`CotC Lines`)|
                                                                        is.na (In.Analysis.Data.6sd.Biz.Generic.Cleaned$`CotC Actual`),
                                                                      NA, In.Analysis.Data.6sd.Biz.Generic.Cleaned$`CotC Achieved %`)




# # 4P Compliance Outlet:  # 4P Compliance Outlet = NA when a) Total Outlet Count = NA | b) Total Outlet Count < # 4P Compliance Outlet
In.Analysis.Data.6sd.Biz.Generic.Cleaned$`# 4P Compliance Outlet` <- ifelse (is.na (In.Analysis.Data.6sd.Biz.Generic.Cleaned$`Total Outlet Count`)|
                                                                                     In.Analysis.Data.6sd.Biz.Generic.Cleaned$`Total Outlet Count`<In.Analysis.Data.6sd.Biz.Generic.Cleaned$`# 4P Compliance Outlet`,
                                                                                   NA, In.Analysis.Data.6sd.Biz.Generic.Cleaned$`# 4P Compliance Outlet`)
# % Perfect Outlet Count:  % Perfect Outlet Count = NA when a)  Total Outlet Count = NA | b) # 4P Compliance Outlet` = NA 
In.Analysis.Data.6sd.Biz.Generic.Cleaned$`% Perfect Outlet Count` <- ifelse (is.na (In.Analysis.Data.6sd.Biz.Generic.Cleaned$`Total Outlet Count`)| 
                                                                      is.na (In.Analysis.Data.6sd.Biz.Generic.Cleaned$`# 4P Compliance Outlet`),
                                                                    NA, In.Analysis.Data.6sd.Biz.Generic.Cleaned$`% Perfect Outlet Count`)




# EB Actual:  EB Actual = NA when a) EB Lines = NA | b) EB Lines <  EB Actual
In.Analysis.Data.6sd.Biz.Generic.Cleaned$`EB Actual` <- ifelse (is.na (In.Analysis.Data.6sd.Biz.Generic.Cleaned$`EB Lines`)|
                                                                  In.Analysis.Data.6sd.Biz.Generic.Cleaned$`EB Lines`< In.Analysis.Data.6sd.Biz.Generic.Cleaned$`EB Actual`,
                                                                NA, In.Analysis.Data.6sd.Biz.Generic.Cleaned$`EB Actual`)
# EB Achieved %:  EB Achieved % = NA when a)  EB Lines = NA | b) EB Actual = NA 
In.Analysis.Data.6sd.Biz.Generic.Cleaned$`EB Achieved %` <- ifelse (is.na (In.Analysis.Data.6sd.Biz.Generic.Cleaned$`EB Lines`)| 
                                                                 is.na (In.Analysis.Data.6sd.Biz.Generic.Cleaned$`EB Actual`),
                                                               NA, In.Analysis.Data.6sd.Biz.Generic.Cleaned$`EB Achieved %`)
# EB - Perfect Outlet Count:  EB-Perfect Oulet Count = NA when a) EB Lines = NA | b) EB Actual = NA
In.Analysis.Data.6sd.Biz.Generic.Cleaned$`EB - Perfect Outlet Count` <- ifelse (is.na (In.Analysis.Data.6sd.Biz.Generic.Cleaned$`EB Lines`)|
                                                                                     is.na (In.Analysis.Data.6sd.Biz.Generic.Cleaned$`EB Actual`),
                                                                                   NA, In.Analysis.Data.6sd.Biz.Generic.Cleaned$`EB - Perfect Outlet Count`)



# RL Actual:  RL Actual = NA when a) Red Lines = NA | b) Red Lines <  RL Actual
In.Analysis.Data.6sd.Biz.Generic.Cleaned$`RL Actual` <- ifelse (is.na (In.Analysis.Data.6sd.Biz.Generic.Cleaned$`Red Lines`)|
                                                                       In.Analysis.Data.6sd.Biz.Generic.Cleaned$`Red Lines`< In.Analysis.Data.6sd.Biz.Generic.Cleaned$`RL Actual`,
                                                                     NA, In.Analysis.Data.6sd.Biz.Generic.Cleaned$`RL Actual`)
# RL Achieved %:  RL Achieved % = NA when a) Red Lines = NA | b) RL Actual = NA
In.Analysis.Data.6sd.Biz.Generic.Cleaned$`RL Achieved %` <- ifelse (is.na (In.Analysis.Data.6sd.Biz.Generic.Cleaned$`Red Lines`)|
                                                                 is.na (In.Analysis.Data.6sd.Biz.Generic.Cleaned$`RL Actual`),
                                                               NA, In.Analysis.Data.6sd.Biz.Generic.Cleaned$`RL Achieved %`)
# Redline-Perfect Oulet Count:  Redline-Perfect Oulet Count = NA when a) Red Lines = NA | b) RL Actual = NA
In.Analysis.Data.6sd.Biz.Generic.Cleaned$`RL - Perfect Outlet Count` <- ifelse (is.na (In.Analysis.Data.6sd.Biz.Generic.Cleaned$`Red Lines`)|
                                                                                    is.na (In.Analysis.Data.6sd.Biz.Generic.Cleaned$`RL Actual`),
                                                                                  NA, In.Analysis.Data.6sd.Biz.Generic.Cleaned$`RL - Perfect Outlet Count`)



# NPD Actual:  NPD Actual = NA when a) NPD Lines  = NA | b) NPD Lines  <  NPD Actual
In.Analysis.Data.6sd.Biz.Generic.Cleaned$`NPD Actual` <- ifelse (is.na (In.Analysis.Data.6sd.Biz.Generic.Cleaned$`NPD Lines`)|
                                                                    In.Analysis.Data.6sd.Biz.Generic.Cleaned$`NPD Lines`< In.Analysis.Data.6sd.Biz.Generic.Cleaned$`NPD Actual`,
                                                                  NA, In.Analysis.Data.6sd.Biz.Generic.Cleaned$`NPD Actual`)
# NPD Achieved %: NPD Achieved % = NA when a) NPD Lines  = NA | b) NPD Actual = NA
In.Analysis.Data.6sd.Biz.Generic.Cleaned$`NPD Achieved %` <- ifelse (is.na (In.Analysis.Data.6sd.Biz.Generic.Cleaned$`NPD Lines`)|
                                                                   is.na (In.Analysis.Data.6sd.Biz.Generic.Cleaned$`NPD Actual`),
                                                                 NA, In.Analysis.Data.6sd.Biz.Generic.Cleaned$`NPD Achieved %`)
# NPD -Perfect Oulet Count:  NPD -Perfect Oulet Count = NA when a) NPD Lines = NA | b) NPD Actual = NA
In.Analysis.Data.6sd.Biz.Generic.Cleaned$`NPD - Perfect Outlet Count` <- ifelse (is.na (In.Analysis.Data.6sd.Biz.Generic.Cleaned$`NPD Lines`)|
                                                                                  is.na (In.Analysis.Data.6sd.Biz.Generic.Cleaned$`NPD Actual`),
                                                                                NA, In.Analysis.Data.6sd.Biz.Generic.Cleaned$`NPD - Perfect Outlet Count`)





# WP Actual:  WP Actual = NA when a) WP Lines = NA | b) WP Lines  <  WP Actual
In.Analysis.Data.6sd.Biz.Generic.Cleaned$`WP Actual` <- ifelse (is.na (In.Analysis.Data.6sd.Biz.Generic.Cleaned$`WP Lines`)|
                                                                    In.Analysis.Data.6sd.Biz.Generic.Cleaned$`WP Lines`< In.Analysis.Data.6sd.Biz.Generic.Cleaned$`WP Actual`,
                                                                  NA, In.Analysis.Data.6sd.Biz.Generic.Cleaned$`WP Actual`)
# WP Achieved %: WP Achieved % = NA when a) WP Lines = NA | b) WP Actual = NA
In.Analysis.Data.6sd.Biz.Generic.Cleaned$`WP Achieved %` <- ifelse (is.na (In.Analysis.Data.6sd.Biz.Generic.Cleaned$`WP Lines`)|
                                                                  is.na (In.Analysis.Data.6sd.Biz.Generic.Cleaned$`WP Actual`),
                                                                NA, In.Analysis.Data.6sd.Biz.Generic.Cleaned$`WP Achieved %`)
# WP - Perfect Outlet Count:  WP - Perfect Outlet Count = NA when a) NPD Lines = NA | b) NPD Actual = NA
In.Analysis.Data.6sd.Biz.Generic.Cleaned$`WP - Perfect Outlet Count` <- ifelse (is.na (In.Analysis.Data.6sd.Biz.Generic.Cleaned$`WP Lines`)|
                                                                                  is.na (In.Analysis.Data.6sd.Biz.Generic.Cleaned$`WP Actual`),
                                                                                NA, In.Analysis.Data.6sd.Biz.Generic.Cleaned$`WP - Perfect Outlet Count`)



In.Analysis.Data.Portfolio.Change.6sd.Biz.Generic.Cleaned <- In.Analysis.Data.6sd.Biz.Generic.Cleaned
rm(In.Analysis.Data.6sd.Biz.Generic.Cleaned)


###### << Ad Hoc >> Export Data to Test Cleansing Steps  ###### ------------------------------
# Just test all the calculations is completed. 
# write.csv (`In.Analysis.Data.Portfolio.Change.6sd.Biz.Generic.Cleaned`, "Temp - In.Analysis.Data.Portfolio.Change.6sd.Biz.Generic.Cleaned.csv") 



###### LOOP > >> Create for Export >> For Regression Model >> Cleaned Model Data  ######
if (Model.Type.Data == "Model.Rerun.KPI") {Cleaned.Model.Rerun.Data <-In.Analysis.Data.Portfolio.Change.6sd.Biz.Generic.Cleaned} else
if (Model.Type.Data == "Model.Refresh.Coefficient") {Cleaned.Model.Refresh.Data <- In.Analysis.Data.Portfolio.Change.6sd.Biz.Generic.Cleaned} 




###### LOOP > Final Regression Input Data Complete Rows * With Country Specific Coding  ###### 
# Subset Out Final Measures used for Regression
# Note Country Specific Code for KPIs not used

RegInput.OmitNA <- subset ( Full.Data, 
                            select = c( "Country Code",                         "Country", 
                                        "Distributor Code",                     "Distributor",    
                                        "Site Code",                            "Site",
                                        "Salesperson ID",                       "Salesperson Name",     
                                        "Month",                                "Training / Testing / Removed Months",
                                        "Portfolio Changes",                    "Included in Analysis",                 
                                        "Salesperson Clustering",               "Appear in all reports",
                                        "GSV",                                  "Salesperson Target",
                                        "% Target - GSV%",                      "# Actual Calls",                     
                                        "Total Working Days",                   "No Of Month End Active Shop in PJP Covered",
                                        "Total Time Spent (Mins)",              "Effective Outlet Time (Mins)", 
                                        "# PJP Complied Calls",                 "% PJP Complied Calls",
                                        "# PJP Complied Days",                  "# Geo Complied Calls",
                                        "% Geo Complied Calls",                 "# Geo Complied Days",
                                        "% ECO to Active Outlet",               "# ECO Achieved",
                                        "% BP to Total Calls",                  "# BP Achieved",
                                        "# PJP Complied Productive Calls",      "Average Daily BP",
                                        "# SKUs(Order)",                        "# SKU's /Actual Calls",                          
                                        "CotC Actual",                          "CotC Achieved %",                      
                                        "# 4P Compliance Outlet",               "% Perfect Outlet Count",
                                        "EB Actual",                            "EB Achieved %",                        
                                        "EB - Perfect Outlet Count",            "RL Actual",
                                        "RL Achieved %",                        "RL - Perfect Outlet Count",
                                        "NPD Actual",                           "NPD Achieved %", 
                                        "NPD - Perfect Outlet Count",           "WP Actual",
                                        "WP Achieved %",                        "WP - Perfect Outlet Count"))

                                        
# ** Country Specific Code
if (unique(RegInput.OmitNA$`Country Code`) == "PH") {
  RegInput.OmitNA <- subset(RegInput.OmitNA, select = -c(`# 4P Compliance Outlet`, `% Perfect Outlet Count`,
                                                         `NPD - Perfect Outlet Count`, 
                                                         `WP - Perfect Outlet Count`))}

if (unique(RegInput.OmitNA$`Country Code`) == "VN") {
  RegInput.OmitNA <- subset(RegInput.OmitNA, select = -c(`# 4P Compliance Outlet`, `% Perfect Outlet Count`,
                                                         `NPD Actual`, `NPD Achieved %`, `NPD - Perfect Outlet Count`, 
                                                         `WP Actual`, `WP Achieved %`, `WP - Perfect Outlet Count`))}




# So as not to exclude rows where GSV = NA
RegInput.OmitNA$GSV <- ifelse (is.na (RegInput.OmitNA$GSV), "NA", RegInput.OmitNA$GSV)

# Get Complete Rows
RegInput.OmitNA <- na.omit(RegInput.OmitNA)





###### LOOP > Final Regression Input Complete Rows w Portfolio CHanges Logic * With Country Specific Coding  ###### 
# Subset Out Final Measures used for Regression and GSV
# Note VN Specific Code for KPIs not used

RegInput.woPortfolio.Change.OmitNA <- subset ( In.Analysis.Data.Portfolio.Change,
                                      select = c( "Country Code",                         "Country", 
                                                  "Distributor Code",                     "Distributor",    
                                                  "Site Code",                            "Site",
                                                  "Salesperson ID",                       "Salesperson Name",     
                                                  "Month",                                "Training / Testing / Removed Months",
                                                  "Portfolio Changes",                    "Included in Analysis",                 
                                                  "Salesperson Clustering",               "Appear in all reports",
                                                  "GSV",                                  "Salesperson Target",
                                                  "% Target - GSV%",                      "# Actual Calls",                     
                                                  "Total Working Days",                   "No Of Month End Active Shop in PJP Covered",
                                                  "Total Time Spent (Mins)",              "Effective Outlet Time (Mins)", 
                                                  "# PJP Complied Calls",                 "% PJP Complied Calls",
                                                  "# PJP Complied Days",                  "# Geo Complied Calls",
                                                  "% Geo Complied Calls",                 "# Geo Complied Days",
                                                  "% ECO to Active Outlet",               "# ECO Achieved",
                                                  "% BP to Total Calls",                  "# BP Achieved",
                                                  "# PJP Complied Productive Calls",      "Average Daily BP",
                                                  "# SKUs(Order)",                        "# SKU's /Actual Calls",                          
                                                  "CotC Actual",                          "CotC Achieved %",                      
                                                  "# 4P Compliance Outlet",               "% Perfect Outlet Count",
                                                  "EB Actual",                            "EB Achieved %",                        
                                                  "EB - Perfect Outlet Count",            "RL Actual",
                                                  "RL Achieved %",                        "RL - Perfect Outlet Count",
                                                  "NPD Actual",                           "NPD Achieved %", 
                                                  "NPD - Perfect Outlet Count",           "WP Actual",
                                                  "WP Achieved %",                        "WP - Perfect Outlet Count"))


# ** Country Specific Code
if (unique(as.character(RegInput.woPortfolio.Change.OmitNA$`Country Code`)) == "PH") {
  RegInput.woPortfolio.Change.OmitNA <- subset(RegInput.woPortfolio.Change.OmitNA, select = -c(`# 4P Compliance Outlet`, `% Perfect Outlet Count`,
                                                                                               `NPD - Perfect Outlet Count`, 
                                                                                               `WP - Perfect Outlet Count`))}

if (unique(RegInput.woPortfolio.Change.OmitNA$`Country Code`) == "VN") {
  RegInput.woPortfolio.Change.OmitNA <- subset(RegInput.woPortfolio.Change.OmitNA, select = -c(`# 4P Compliance Outlet`, `% Perfect Outlet Count`,
                                                                                               `NPD Actual`, `NPD Achieved %`, `NPD - Perfect Outlet Count`,
                                                                                               `WP Actual`, `WP Achieved %`, `WP - Perfect Outlet Count`))}



# So as not to exclude rows where GSV = NA
RegInput.woPortfolio.Change.OmitNA$GSV <- ifelse (is.na (RegInput.woPortfolio.Change.OmitNA$GSV),
                                                  "NA", RegInput.woPortfolio.Change.OmitNA$GSV)

# Get Complete Rows
RegInput.woPortfolio.Change.OmitNA <- na.omit(RegInput.woPortfolio.Change.OmitNA)
rm(In.Analysis.Data.Portfolio.Change)



###### LOOP > Final Regression Input Complete Rows w Portfolio CHanges Logic within 6sd Data Uncleaned* With Country Specific Coding  ###### 
# Subset Out Final Measures used for Regression
# Note VN Specific Code for KPIs not used
RegInput.6sd.Uncleaned.OmitNA <- subset ( In.Analysis.Data.Portfolio.Change.6sd.Uncleaned, 
                                          select = c( "Country Code",                         "Country", 
                                                      "Distributor Code",                     "Distributor",    
                                                      "Site Code",                            "Site",
                                                      "Salesperson ID",                       "Salesperson Name",     
                                                      "Month",                                "Training / Testing / Removed Months",
                                                      "Portfolio Changes",                    "Included in Analysis",                 
                                                      "Salesperson Clustering",               "Appear in all reports",
                                                      "GSV",                                  "Salesperson Target",
                                                      "% Target - GSV%",                      "# Actual Calls",                     
                                                      "Total Working Days",                   "No Of Month End Active Shop in PJP Covered",
                                                      "Total Time Spent (Mins)",              "Effective Outlet Time (Mins)", 
                                                      "# PJP Complied Calls",                 "% PJP Complied Calls",
                                                      "# PJP Complied Days",                  "# Geo Complied Calls",
                                                      "% Geo Complied Calls",                 "# Geo Complied Days",
                                                      "% ECO to Active Outlet",               "# ECO Achieved",
                                                      "% BP to Total Calls",                  "# BP Achieved",
                                                      "# PJP Complied Productive Calls",      "Average Daily BP",
                                                      "# SKUs(Order)",                        "# SKU's /Actual Calls",                          
                                                      "CotC Actual",                          "CotC Achieved %",                      
                                                      "# 4P Compliance Outlet",               "% Perfect Outlet Count",
                                                      "EB Actual",                            "EB Achieved %",                        
                                                      "EB - Perfect Outlet Count",            "RL Actual",
                                                      "RL Achieved %",                        "RL - Perfect Outlet Count",
                                                      "NPD Actual",                           "NPD Achieved %", 
                                                      "NPD - Perfect Outlet Count",           "WP Actual",
                                                      "WP Achieved %",                        "WP - Perfect Outlet Count"))

# ** Country Specific Code
if (unique(RegInput.6sd.Uncleaned.OmitNA$`Country Code`) == "PH") {
  RegInput.6sd.Uncleaned.OmitNA <- subset(RegInput.6sd.Uncleaned.OmitNA, select = -c(`# 4P Compliance Outlet`, `% Perfect Outlet Count`,
                                                                                     `NPD - Perfect Outlet Count`, 
                                                                                     `WP - Perfect Outlet Count`))}

if (unique(RegInput.6sd.Uncleaned.OmitNA$`Country Code`) == "VN") {
  RegInput.6sd.Uncleaned.OmitNA <- subset(RegInput.6sd.Uncleaned.OmitNA, select = -c(`# 4P Compliance Outlet`, `% Perfect Outlet Count`,
                                                                                     `NPD Actual`, `NPD Achieved %`, `NPD - Perfect Outlet Count`, 
                                                                                     `WP Actual`, `WP Achieved %`, `WP - Perfect Outlet Count`))}


# So as not to exclude rows where GSV = NA
RegInput.6sd.Uncleaned.OmitNA$GSV <- ifelse (is.na (RegInput.6sd.Uncleaned.OmitNA$GSV),
                                                    "NA", RegInput.6sd.Uncleaned.OmitNA$GSV)

# Get Complete Rows
RegInput.6sd.Uncleaned.OmitNA <- na.omit(RegInput.6sd.Uncleaned.OmitNA)
rm(In.Analysis.Data.Portfolio.Change.6sd.Uncleaned)



###### LOOP > Final Regression Input Complete Rows w Portfolio CHanges Logic within 6sd DataCleaned with Business Logic * With Country Specific Coding  ###### 
# Subset Out Final Measures used for Regression and GSV
# Note VN Specific Code for KPIs not used

RegInput.6sd.Biz.Cleaned.OmitNA <- subset ( In.Analysis.Data.Portfolio.Change.6sd.Biz.Cleaned, 
                                            select = c( "Country Code",                         "Country", 
                                                        "Distributor Code",                     "Distributor",    
                                                        "Site Code",                            "Site",
                                                        "Salesperson ID",                       "Salesperson Name",     
                                                        "Month",                                "Training / Testing / Removed Months",
                                                        "Portfolio Changes",                    "Included in Analysis",                 
                                                        "Salesperson Clustering",               "Appear in all reports",
                                                        "GSV",                                  "Salesperson Target",
                                                        "% Target - GSV%",                      "# Actual Calls",                     
                                                        "Total Working Days",                   "No Of Month End Active Shop in PJP Covered",
                                                        "Total Time Spent (Mins)",              "Effective Outlet Time (Mins)", 
                                                        "# PJP Complied Calls",                 "% PJP Complied Calls",
                                                        "# PJP Complied Days",                  "# Geo Complied Calls",
                                                        "% Geo Complied Calls",                 "# Geo Complied Days",
                                                        "% ECO to Active Outlet",               "# ECO Achieved",
                                                        "% BP to Total Calls",                  "# BP Achieved",
                                                        "# PJP Complied Productive Calls",      "Average Daily BP",
                                                        "# SKUs(Order)",                        "# SKU's /Actual Calls",                          
                                                        "CotC Actual",                          "CotC Achieved %",                      
                                                        "# 4P Compliance Outlet",               "% Perfect Outlet Count",
                                                        "EB Actual",                            "EB Achieved %",                        
                                                        "EB - Perfect Outlet Count",            "RL Actual",
                                                        "RL Achieved %",                        "RL - Perfect Outlet Count",
                                                        "NPD Actual",                           "NPD Achieved %", 
                                                        "NPD - Perfect Outlet Count",           "WP Actual",
                                                        "WP Achieved %",                        "WP - Perfect Outlet Count"))


# ** Country Specific Code
if (unique(RegInput.6sd.Biz.Cleaned.OmitNA$`Country Code`) == "PH") {
  RegInput.6sd.Biz.Cleaned.OmitNA <- subset(RegInput.6sd.Biz.Cleaned.OmitNA, select = -c(`# 4P Compliance Outlet`, `% Perfect Outlet Count`,
                                                                                         `NPD - Perfect Outlet Count`, 
                                                                                         `WP - Perfect Outlet Count`))}

if (unique(RegInput.6sd.Biz.Cleaned.OmitNA$`Country Code`) == "VN") {
  RegInput.6sd.Biz.Cleaned.OmitNA <- subset(RegInput.6sd.Biz.Cleaned.OmitNA, select = -c(`# 4P Compliance Outlet`, `% Perfect Outlet Count`,
                                                                                         `NPD Actual`, `NPD Achieved %`, `NPD - Perfect Outlet Count`, 
                                                                                         `WP Actual`, `WP Achieved %`, `WP - Perfect Outlet Count`))}

# So as not to exclude rows where GSV = NA
RegInput.6sd.Biz.Cleaned.OmitNA$GSV <- ifelse (is.na (RegInput.6sd.Biz.Cleaned.OmitNA$GSV),
                                               "NA", RegInput.6sd.Biz.Cleaned.OmitNA$GSV)

# Get Complete Rows
RegInput.6sd.Biz.Cleaned.OmitNA <- na.omit(RegInput.6sd.Biz.Cleaned.OmitNA)
rm(In.Analysis.Data.Portfolio.Change.6sd.Biz.Cleaned)



###### LOOP > Final Regression Input Complete Rows w Portfolio CHanges Logic within 6sd DataCleaned with Business and Generic Logic * With Country Specific Coding  ###### 
# Subset Out Final Measures used for Regression and GSV
# Note VN Specific Code for KPIs not used

RegInput.6sd.Biz.Generic.Cleaned.OmitNA <- subset ( In.Analysis.Data.Portfolio.Change.6sd.Biz.Generic.Cleaned, 
                                                    select = c( "Country Code",                         "Country", 
                                                                "Distributor Code",                     "Distributor",    
                                                                "Site Code",                            "Site",
                                                                "Salesperson ID",                       "Salesperson Name",     
                                                                "Month",                                "Training / Testing / Removed Months",
                                                                "Portfolio Changes",                    "Included in Analysis",                 
                                                                "Salesperson Clustering",               "Appear in all reports",
                                                                "GSV",                                  "Salesperson Target",
                                                                "% Target - GSV%",                      "# Actual Calls",                     
                                                                "Total Working Days",                   "No Of Month End Active Shop in PJP Covered",
                                                                "Total Time Spent (Mins)",              "Effective Outlet Time (Mins)", 
                                                                "# PJP Complied Calls",                 "% PJP Complied Calls",
                                                                "# PJP Complied Days",                  "# Geo Complied Calls",
                                                                "% Geo Complied Calls",                 "# Geo Complied Days",
                                                                "% ECO to Active Outlet",               "# ECO Achieved",
                                                                "% BP to Total Calls",                  "# BP Achieved",
                                                                "# PJP Complied Productive Calls",      "Average Daily BP",
                                                                "# SKUs(Order)",                        "# SKU's /Actual Calls",                          
                                                                "CotC Actual",                          "CotC Achieved %",                      
                                                                "# 4P Compliance Outlet",               "% Perfect Outlet Count",
                                                                "EB Actual",                            "EB Achieved %",                        
                                                                "EB - Perfect Outlet Count",            "RL Actual",
                                                                "RL Achieved %",                        "RL - Perfect Outlet Count",
                                                                "NPD Actual",                           "NPD Achieved %", 
                                                                "NPD - Perfect Outlet Count",           "WP Actual",
                                                                "WP Achieved %",                        "WP - Perfect Outlet Count"))


# ** Country Specific Code
if (unique(RegInput.6sd.Biz.Generic.Cleaned.OmitNA$`Country Code`) == "PH") {
  RegInput.6sd.Biz.Generic.Cleaned.OmitNA <- subset(RegInput.6sd.Biz.Generic.Cleaned.OmitNA, select = -c(`# 4P Compliance Outlet`, `% Perfect Outlet Count`,
                                                                                                         `NPD - Perfect Outlet Count`, 
                                                                                                         `WP - Perfect Outlet Count`))}

if (unique(RegInput.6sd.Biz.Generic.Cleaned.OmitNA$`Country Code`) == "VN") {
  RegInput.6sd.Biz.Generic.Cleaned.OmitNA <- subset(RegInput.6sd.Biz.Generic.Cleaned.OmitNA, select = -c(`# 4P Compliance Outlet`, `% Perfect Outlet Count`,
                                                                                                         `NPD Actual`, `NPD Achieved %`, `NPD - Perfect Outlet Count`, 
                                                                                                         `WP Actual`, `WP Achieved %`, `WP - Perfect Outlet Count`))}


# So as not to exclude rows where GSV = NA
RegInput.6sd.Biz.Generic.Cleaned.OmitNA$GSV <- ifelse (is.na (RegInput.6sd.Biz.Generic.Cleaned.OmitNA$GSV),
                                                       "NA", RegInput.6sd.Biz.Generic.Cleaned.OmitNA$GSV)

# Get Complete Rows
RegInput.6sd.Biz.Generic.Cleaned.OmitNA <- na.omit(RegInput.6sd.Biz.Generic.Cleaned.OmitNA)
rm(In.Analysis.Data.Portfolio.Change.6sd.Biz.Generic.Cleaned)





###### LOOP > >> Standard >> May Need Revision >> Get List of Data Frames ######
List.Data <- list()
  List.Data$Full.Data <- Full.Data
  List.Data$Full.Data.GSV <- Full.Data.GSV
  List.Data$Full.Data.In.All.Reports <- Full.Data.In.All.Reports
  List.Data$RegInput.OmitNA <- RegInput.OmitNA
  List.Data$RegInput.woPortfolio.Change.OmitNA <- RegInput.woPortfolio.Change.OmitNA
  List.Data$RegInput.6sd.Uncleaned.OmitNA <- RegInput.6sd.Uncleaned.OmitNA
  List.Data$RegInput.6sd.Biz.Cleaned.OmitNA <- RegInput.6sd.Biz.Cleaned.OmitNA
  List.Data$RegInput.6sd.Biz.Generic.Cleaned.OmitNA <- RegInput.6sd.Biz.Generic.Cleaned.OmitNA

  
###### LOOP LOOP START >> Sum GSV and Count Rows For All ######
# Run and Save Table (Note Manual Shortening of Naming) 
  
Final.Biz.Rep <- NULL

system.time (for ( Selected.DataSet in names(List.Data)) {
  # Selected.DataSet = "Full.Data" 
  # Selected.DataSet = "Full.Data.GSV" 
  # Selected.DataSet = "Full.Data.In.All.Reports" 
  # Selected.DataSet = "RegInput.OmitNA" 
  # Selected.DataSet = "RegInput.woPortfolio.Change.OmitNA" 
  # Selected.DataSet = "RegInput.6sd.Uncleaned.OmitNA" 
  # Selected.DataSet = "RegInput.6sd.Biz.Cleaned.OmitNA" 
  # Selected.DataSet = "RegInput.6sd.Biz.Generic.Cleaned.OmitNA" 
  
  Data.Name = ifelse (Selected.DataSet == "Full.Data","Full.Data",
                      ifelse (Selected.DataSet == "Full.Data.GSV","Full.Data.GSV",
                              ifelse (Selected.DataSet == "Full.Data.In.All.Reports","Full.Data.In.All.Reports",
                                      ifelse (Selected.DataSet == "RegInput.OmitNA", "RegInput.OmitNA",
                                              ifelse (Selected.DataSet == "RegInput.woPortfolio.Change.OmitNA","RegInput.woPortfolio.Change.OmitNA",
                                                      ifelse (Selected.DataSet == "RegInput.6sd.Uncleaned.OmitNA","RegInput.6sd.Uncleaned.OmitNA",
                                                              ifelse (Selected.DataSet == "RegInput.6sd.Biz.Cleaned.OmitNA","RegInput.6sd.Biz.Cleaned.OmitNA",
                                                                      ifelse (Selected.DataSet == "RegInput.6sd.Biz.Generic.Cleaned.OmitNA","RegInput.6sd.Biz.Generic.Cleaned.OmitNA",
                                                                              "ERROR"))))))))
      #print (Selected.DataSet) 
      #print (Data.Name)
  
  df.Final <- data.frame(List.Data[Selected.DataSet],check.names = FALSE)
  names(df.Final) <- substring(names(df.Final),nchar(Selected.DataSet)+2)
  df.Final$`Salesperson Clustering` <- addNA(df.Final$`Salesperson Clustering`)
  df.Final$GSV <- as.numeric (df.Final$GSV)
  df.Final$Month <- as.character(df.Final$Month)
  df.Final$Count = 1
  
  Agg.GSV  <- aggregate (GSV ~ `Salesperson Clustering` + `Training / Testing / Removed Months` + Month, data = df.Final, sum, na.rm = T)
  Agg.GSV$`GSV / Lines` = "GSV"
  names(Agg.GSV)[4] <- "Value"
      print ("test GSV Aggregation")
      print (sum (df.Final$GSV, na.rm = T))
      print (sum (Agg.GSV$Value))
  
  Agg.Lines  <- aggregate ( Count  ~ `Salesperson Clustering` + `Training / Testing / Removed Months` + Month, data = df.Final, sum, na.rm = T)
  Agg.Lines$`GSV / Lines` = "Lines"
  names(Agg.Lines)[4] <- "Value"
      print ("test Count Aggregation")
      print (sum (df.Final$Count, na.rm = T))
      print (sum (Agg.Lines$Value))
      
  Agg.Full.Data  <- rbind (Agg.GSV, Agg.Lines)
  Agg.Full.Data$Data = Data.Name
  
  Final.Biz.Rep <-  rbind ( Final.Biz.Rep , Agg.Full.Data)
                            
  rm(Selected.DataSet)
  rm(df.Final)
  rm(Agg.Lines)
  rm(Agg.GSV)
  rm(Agg.Full.Data)
  rm(Data.Name)
})

###### LOOP LOOP END >> Sum GSV and Count Rows For All ######



###### LOOP > Identify if the Final Biz Rep Data is for Model Rerun or Refresh ######

Final.Biz.Rep$`For Model Rerun / Refresh` <- ifelse(Model.Type.Data == "Model.Rerun.KPI", "For Model Rerun (KPI Refresh)",
                                                    ifelse (Model.Type.Data == "Model.Refresh.Coefficient", "For Model Refresh (Coefficient Refresh)",
                                                            "ERROR"))

if (Model.Type.Data == "Model.Rerun.KPI") {Final.Biz.Rep.Model.Rerun <-Final.Biz.Rep} else
  if (Model.Type.Data == "Model.Refresh.Coefficient") {Final.Biz.Rep.Model.Refresh <- Final.Biz.Rep} 

rm(Final.Biz.Rep)

}

###### LOOP END > Create Loop for Clean Data For Modelling and Data Cleansing Biz Rep  ######

rm(Full.Data)
rm(Full.Data.GSV)
rm(Full.Data.In.All.Reports)
rm(RegInput.OmitNA)
rm(RegInput.woPortfolio.Change.OmitNA)
rm(RegInput.6sd.Uncleaned.OmitNA)
rm(RegInput.6sd.Biz.Cleaned.OmitNA) 
rm(RegInput.6sd.Biz.Generic.Cleaned.OmitNA)


###### >> Create For Export >> For Data Cleansing Report >> Data Cleansing Biz Rep ######    

Final.Biz.Rep.All <- rbind(Final.Biz.Rep.Model.Rerun, Final.Biz.Rep.Model.Refresh)

rm(Final.Biz.Rep.Model.Rerun)
rm(Final.Biz.Rep.Model.Refresh)







###### << Manual>> write the Portfolio Changes KPI Refresh Report to correct folder ######

# For PH
#write.csv (Portfolio.Changes.KPI.Refresh, "//falmumapp39/Unilever Singapore (15-AIM-1275)/SIO Analytics/PH/Data/06 Dec 2018/All data/Portfolio Changes KPI Refresh Report.csv", row.names=FALSE, fileEncoding="UTF-8")

# For VN 
#write.csv (Portfolio.Changes.KPI.Refresh, "C:/Users/huishan.chin/OneDrive - Unilever/1. LE SIO/1. Working Folder/2. VN/2. Final Output/1. VN Data Cleaning Representation - 180207/Portfolio Changes KPI Refresh Report.csv", row.names=FALSE, fileEncoding="UTF-8") 





###### << Manual>> write the Portfolio Changes Coefficient Refresh Report to correct folder ######

# For PH
#write.csv (Portfolio.Changes.Coefficient.Refresh, "//falmumapp39/Unilever Singapore (15-AIM-1275)/SIO Analytics/PH/Data/06 Dec 2018/All data/Portfolio Changes Coefficient Refresh Report.csv", row.names=FALSE, fileEncoding="UTF-8")

# For VN 
#write.csv (Portfolio.Changes.Coefficient.Refresh, "C:/Users/huishan.chin/OneDrive - Unilever/1. LE SIO/1. Working Folder/2. VN/2. Final Output/1. VN Data Cleaning Representation - 180207/Portfolio Changes Coefficient Refresh Report.csv", row.names=FALSE, fileEncoding="UTF-8")





###### << Manual>> write the Extreme Outliers Report to correct folder ######

# For PH
#write.xlsx (Report.Extreme.Outliers, "//falmumapp39/Unilever Singapore (15-AIM-1275)/SIO Analytics/PH/Data/06 Dec 2018/All data/Extreme Outliers Report.xlsx") 

# For VN 
#write.xlsx (Report.Extreme.Outliers, "C:/Users/huishan.chin/OneDrive - Unilever/1. LE SIO/1. Working Folder/2. VN/2. Final Output/1. VN Data Cleaning Representation - 180207/Extreme Outliers Report.xlsx") 





###### << Manual>> write the Cleaned Model Rerun Data to the correct folder ######
# For PH
write.csv (Cleaned.Model.Rerun.Data, paste0("//falmumapp39/RTM Vietnam (19-AIM-3081)/SIO Analytics/PH/Data/03 Jan 2020/All Data/PH Cleaned Model Rerun Data - 060120_",z,".csv"), row.names=FALSE, fileEncoding="UTF-8") 

# For VN 
#write.csv (Cleaned.Model.Rerun.Data, "C:/Users/huishan.chin/OneDrive - Unilever/1. LE SIO/1. Working Folder/2. VN/2. Final Output/VN Cleaned Model Rerun Data - 180207.csv", row.names=FALSE, fileEncoding="UTF-8")


}


###### << Manual>> write the Cleaned Model Refresh Data to the correct folder ######

# For PH
#write.csv (Cleaned.Model.Refresh.Data, "//falmumapp39/Unilever Singapore (15-AIM-1275)/SIO Analytics/PH/Data/06 Dec 2018/All data/PH Cleaned Model Refresh Data - 180917.csv", row.names=FALSE, fileEncoding="UTF-8")

# For VN 
# write.csv (Cleaned.Model.Refresh.Data, "C:/Users/huishan.chin/OneDrive - Unilever/1. LE SIO/1. Working Folder/2. VN/2. Final Output/VN Cleaned Model Refresh Data - 180207.csv", row.names=FALSE, fileEncoding="UTF-8")





###### << Manual>> write the Data Cleansing Repretation Data to the correct folder ######

# For PH
#write.csv (Final.Biz.Rep.All, "C:/Users/huishan.chin/OneDrive - Unilever/1. LE SIO/1. Working Folder/1. PH/2. Final Output/1. PH Data Cleaning Representation - 180917/Final Data Representation.csv", row.names=FALSE, fileEncoding="UTF-8")

# For VN 
#write.csv (Final.Biz.Rep.All, "C:/Users/huishan.chin/OneDrive - Unilever/1. LE SIO/1. Working Folder/2. VN/2. Final Output/1. VN Data Cleaning Representation - 180207/Final Data Representation.csv", row.names=FALSE, fileEncoding="UTF-8")




