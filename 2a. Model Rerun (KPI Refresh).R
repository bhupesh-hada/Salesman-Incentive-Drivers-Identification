rm(list = ls())
library(dplyr)
library(data.table)
library(gtools)
library(xlsx)
library(bit64)
library(openxlsx)
library(tidyr)
###### << Manual>> Set Working Directory where all files will be saved ######

# For PH
setwd ("//falmumapp39/RTM Vietnam (19-AIM-3081)/SIO Analytics/PH/Data/03 Jan 2020/All Data/")

# For VN 
# setwd ("C:/Users/huishan.chin/OneDrive - Unilever/1. LE SIO/1. Working Folder/2. VN/2. Final Output/2a. VN Model Rerun (KPI Refresh) - 180207/")



###### << Manual>> Upload CSV ######

# For PH
Cleaned.Model.Rerun.Data <- read.csv( "//falmumapp39/RTM Vietnam (19-AIM-3081)/SIO Analytics/PH/Data/03 Jan 2020/All Data/PH Cleaned Model Rerun Data - 060120_15.csv",
                               check.names = FALSE, encoding = "UTF-8")

# For VN
#Cleaned.Model.Rerun.Data <- read.csv("C:/Users/huishan.chin/OneDrive - Unilever/1. LE SIO/1. Working Folder/2. VN/2. Final Output/VN Cleaned Model Rerun Data - 171203.csv",
#                              check.names = FALSE, encoding = "UTF-8")





###### << Manual>> Select the Months ######
# Need to have a value for all the following (No NAs)

## Start and End Month Needs to Emcompass all Months used 
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



###### Select the KPI Input for Stepwise.Input * w Country Specific Coding  ######

Final.Data.Stepwise.KPI <- Cleaned.Model.Rerun.Data
Final.Data.Stepwise.KPI <- subset (Final.Data.Stepwise.KPI, 
                              select = c( "Salesperson Clustering",              "Month",
                                          "Training / Testing / Removed Months", 
                                          "Salesperson Target",
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




# ** Country Specific Code: Need to Remove Variables that are not used for the country
if (unique(Cleaned.Model.Rerun.Data$`Country Code`) == "PH") {
  Final.Data.Stepwise.KPI <- subset(Final.Data.Stepwise.KPI, select = -c(`# 4P Compliance Outlet`, `% Perfect Outlet Count`,
                                                                           `NPD - Perfect Outlet Count`, 
                                                                           `WP - Perfect Outlet Count`))}

if (unique(Cleaned.Model.Rerun.Data$`Country Code`) == "VN") {
  Final.Data.Stepwise.KPI <- subset(Final.Data.Stepwise.KPI, select = -c(`# 4P Compliance Outlet`, `% Perfect Outlet Count`,
                                                                           `NPD Actual`, `NPD Achieved %`, `NPD - Perfect Outlet Count`, 
                                                                           `WP Actual`, `WP Achieved %`, `WP - Perfect Outlet Count`))}



######## List of Business Constraints and KPIs with no Constraints  * w Country Specific Coding ######## 
# Either / Or within a constraint unless otherwise indicated (which will be split up into a/b)
# Ensure that all new constraints starts with "Constraints." then xx
# -- and that no other data has a name that contains "Constraints." anywhere else

Constraints.05 = c("Total Time Spent (Mins)", "Effective Outlet Time (Mins)")
Constraints.06 = c("% PJP Complied Calls", "# PJP Complied Calls", "# PJP Complied Days")
Constraints.07 = c("% Geo Complied Calls", "# Geo Complied Calls", "# Geo Complied Days")
Constraints.08 = c("% ECO to Active Outlet", "# ECO Achieved")
Constraints.09a = c("% BP to Total Calls", "# BP Achieved", "# PJP Complied Productive Calls", "Average Daily BP")
## Only if Constraint.9a = "% BP to total calls (Cal.)" then Constraint.9b, else exclude Constraint.9b
## Hard Coded Below
Constraints.09b = c("# Actual Calls") 
Constraints.10 = c("# SKUs(Order)", "# SKU's /Actual Calls") 
Constraints.11 = c("CotC Achieved %", "CotC Actual") 
Constraints.12 = c("% Perfect Outlet Count", "# 4P Compliance Outlet", "Empty") 
## Only if Constraint.12 = NULL then Constraint.13-16, else exclude Constraint.
## Hard Coded Below
Constraints.13 = c("EB Achieved %", "EB Actual", "EB - Perfect Outlet Count") 
Constraints.14 = c("RL Achieved %", "RL Actual", "RL - Perfect Outlet Count") 
Constraints.15 = c("NPD Achieved %", "NPD Actual", "NPD - Perfect Outlet Count") 
Constraints.16 = c("WP Achieved %", "WP Actual", "WP - Perfect Outlet Count") 


# ** Country Specific Code: Need to Remove Variables that are not used for the country

if (unique(Cleaned.Model.Rerun.Data$`Country Code`) == "PH") {
  Constraints.12 <- Constraints.12[!is.element(Constraints.12, c("# 4P Compliance Outlet"))]
  Constraints.12 <- Constraints.12[!is.element(Constraints.12, c("% Perfect Outlet Count"))]
  Constraints.15 <- Constraints.15[!is.element(Constraints.15, c("NPD - Perfect Outlet Count"))]
  Constraints.16 <- Constraints.16[!is.element(Constraints.16, c("WP - Perfect Outlet Count"))]
}

if (unique(Cleaned.Model.Rerun.Data$`Country Code`) == "VN") {
  Constraints.12 <- Constraints.12[!is.element(Constraints.12, c("# 4P Compliance Outlet"))]
  Constraints.12 <- Constraints.12[!is.element(Constraints.12, c("% Perfect Outlet Count"))]
  Constraints.15 <- Constraints.15[!is.element(Constraints.15, c("NPD Achieved %"))]
  Constraints.15 <- Constraints.15[!is.element(Constraints.15, c("NPD Actual"))]
  Constraints.15 <- Constraints.15[!is.element(Constraints.15, c("NPD - Perfect Outlet Count"))]
  Constraints.16 <- Constraints.16[!is.element(Constraints.16, c("WP Achieved %"))]
  Constraints.16 <- Constraints.16[!is.element(Constraints.16, c("WP Actual"))]
  Constraints.16 <- Constraints.16[!is.element(Constraints.16, c("WP - Perfect Outlet Count"))]
}




######## Different Combination of Data for Regression (Non Either/Or Business Constraint hard coded) ######## 

# List of Constraints
List.Constraint <- list()
for ( Constraint in ls (pattern = "Constraints.")) {  
  List.Constraint[[Constraint]] <- get(Constraint)
  rm (Constraint)}

# List of KPIs that have no constraints (will appear in all combinations)
List.No.Constraint <- subset (Final.Data.Stepwise.KPI, select = -c(`Salesperson Clustering` , `Month`, `Training / Testing / Removed Months`))
List.No.Constraint = names(List.No.Constraint)
for (Constraint in List.Constraint) { 
  List.No.Constraint = List.No.Constraint[!List.No.Constraint %in% Constraint] 
  rm (Constraint)}

# Create dataframe of all possible combinations of constraints
List.of.Combination <- expand.grid(List.Constraint)

# Hard Coded Constraint
# Need to convert all to characters to recognise specific string
List.of.Combination <- data.frame(lapply(List.of.Combination, as.character), stringsAsFactors=FALSE)
List.of.Combination$Constraints.09b <- ifelse(List.of.Combination$Constraints.09a == "% BP to Total Calls", 
                                              List.of.Combination$Constraints.09b, NA)
List.of.Combination$Constraints.12 <- ifelse(List.of.Combination$Constraints.12== "Empty", NA, List.of.Combination$Constraints.12)
List.of.Combination$Constraints.13 <- ifelse(is.na(List.of.Combination$Constraints.12), List.of.Combination$Constraints.13, NA)
List.of.Combination$Constraints.14 <- ifelse(is.na(List.of.Combination$Constraints.12), List.of.Combination$Constraints.14, NA)
List.of.Combination$Constraints.15 <- ifelse(is.na(List.of.Combination$Constraints.12), List.of.Combination$Constraints.15, NA)
List.of.Combination$Constraints.16 <- ifelse(is.na(List.of.Combination$Constraints.12), List.of.Combination$Constraints.16, NA)


# Add in list of KPIs with no constraints across all combinations
# Rearrange Columns
# Print out List of Combinations for Reference
for (No.Constraint.KPI in List.No.Constraint) {
  List.of.Combination <- cbind (No.Constraint.KPI = No.Constraint.KPI, List.of.Combination)
  colnames(List.of.Combination)[colnames(List.of.Combination)=="No.Constraint.KPI"] <- No.Constraint.KPI 
  rm(No.Constraint.KPI)}
List.of.Combination <- cbind (List.of.Combination[c(List.No.Constraint)],
                              List.of.Combination[c(ls (pattern = "Constraints."))])

  
###### Export >> Write the List of Combinations to working Directory ######

List.of.Combination <- unique(List.of.Combination)
write.csv (List.of.Combination, "2a. Stepwise List of Combination_15.csv") 





###### Create List for Loop of the different combinations ######
# Make sure to keep string as characters instead of factors 
# -- This will affect when pulling the combination from the training set later
List.of.Combination <- data.frame(lapply(List.of.Combination, as.character), stringsAsFactors=FALSE)
List.of.Combination <- t(List.of.Combination)
rownames(List.of.Combination) <- NULL      
List.of.Combination <- as.list (as.data.frame (List.of.Combination, stringsAsFactors= FALSE))

# Get rid of NAs
List.of.Combination <- lapply(List.of.Combination, function(x) x[!is.na(x)])

rm(list = ls (pattern = "Constraints."))
rm(List.Constraint)
rm(List.No.Constraint)





######## (6hrs+) LOOP START >> All Clusters - Finding the Best Combination of Inputs for Stepwise Regression (Run Once then skip) ########

Salesperson.Clustering <- as.character(na.omit(as.character(unique(Final.Data.Stepwise.KPI$`Salesperson Clustering`))))
#Salesperson.Clustering <- Salesperson.Clustering[1:3]
Final.Stepwise.Combination.by.Cluster <- NULL

system.time ( for (Selected.Cluster in Salesperson.Clustering) { 
  # Selected.Cluster = "Drug Store Specialist Sales Route Type"
  # Selected.Cluster = "General Trade Account Specialist Sales Route Type"
  # Selected.Cluster = "Pre-Selling Public Market and Multi Channel Sales Route Type"
  # Selected.Cluster = "Pre-Selling Sari Sari Stores Sales Route Type"
  # Selected.Cluster = "Van Public Market and Multi Channel Sales Route Type"
  # Selected.Cluster = "Van Sari Sari Stores Sales Route Type"
  
  # Selected.Cluster = "HCF Product Group"
  # Selected.Cluster = "MIX Product Group"
  # Selected.Cluster = "PC Product Group"
  # Selected.Cluster = "SPW Product Group"
  
  print (Selected.Cluster)
    
  
  ###### (6hrs+) LOOP > Create Stepwise Regression Data Sets for Selected Cluster ######
  Stepwise.Input.Cluster <-  subset ( Final.Data.Stepwise.KPI, `Salesperson Clustering` == Selected.Cluster)
  
  # Only Complete Rows
  Stepwise.Input.Cluster.OmitNA <- na.omit(Stepwise.Input.Cluster)
  
  # Subset only Training Months - Measures Only
  Training.Set <- subset (Stepwise.Input.Cluster.OmitNA, `Training / Testing / Removed Months` == "KPI Refresh - Training Months")
  Training.Set <- subset (Training.Set, select = -c(`Salesperson Clustering` , `Month`, `Training / Testing / Removed Months`))
  
  # Subset only Testing Months - Measures Only
  Testing.Set.Full.Range <- subset (Stepwise.Input.Cluster.OmitNA, `Training / Testing / Removed Months` == "KPI Refresh - Testing Months")
  Testing.Set.Full.Range <- subset (Testing.Set.Full.Range, select = -c(`Salesperson Clustering`, `Month`, `Training / Testing / Removed Months`))
  
  # Subset rows in testing months where independent variables are within Range of tesing Set
  Measures.Names <- names (subset (Training.Set, select = -c(`% Target - GSV%`)))
  Testing.Set.Within.Range <- Testing.Set.Full.Range
  
  for (Measure in Measures.Names) {
    max.range <-  max (`Training.Set`[which ( colnames (`Training.Set`) ==  Measure )])
    min.range <-  min (`Training.Set`[which ( colnames (`Training.Set`) ==  Measure )])
    `Testing.Set.Within.Range`[which ( colnames (`Testing.Set.Within.Range`) ==  Measure )] <- sapply (
      (`Testing.Set.Within.Range`[which ( colnames (`Testing.Set.Within.Range`) ==  Measure )]), function (x) {
        ifelse (x < min.range | x > max.range, NA, x)})
    rm(Measure)
    rm(min.range)
    rm(max.range)
  }
  
  rm(Stepwise.Input.Cluster)
  rm(Stepwise.Input.Cluster.OmitNA)
  rm(Measures.Names)
  

  
  ###### (6hrs+) LOOP LOOP Start >> Getting All Combination outputs (Stepwise Regression Looping) for Selected Cluster ######

  
  Stepwise.Combination.Summary <- NULL
  
  system.time ( for ( Combination in 1:length(List.of.Combination)) {
    
    print (Combination)
    Stepwise.Combination.Input <- subset ( Training.Set, select = c(List.of.Combination[[Combination]])) 
    
    library(MASS)
    Combination.fit  <- lm (`% Target - GSV%` ~ . ,  data = `Stepwise.Combination.Input`)
    Combination.step <- stepAIC(Combination.fit, direction="both",trace=0) 
    # Combination.step$anova 
    # colnames(Combination.step$model) 
    Combination.Data <- Combination.step$model
    Combination.Final <-lm( `% Target - GSV%` ~., data = Combination.Data)
    # summary (Combination.Final)
    Combination.adj.r.squared <- summary (Combination.Final)$adj.r.squared
    Combination.p.value <-  pf(summary (Combination.Final)$fstatistic[1],
                               summary (Combination.Final)$fstatistic[2],
                               summary (Combination.Final)$fstatistic[3],
                               lower.tail = FALSE )
    
    # Rbind the cluster data to the final data set
    `Stepwise.Combination.Summary` <- rbind(Stepwise.Combination.Summary,
                                            data.frame (Combination, Combination.adj.r.squared, Combination.p.value))
   
    rm(Combination)
    rm(Stepwise.Combination.Input)
    rm(Combination.fit)
    rm(Combination.step)
    rm(Combination.Data)
    rm(Combination.Final)
    rm(Combination.adj.r.squared)
    rm(Combination.p.value)
    
  })
  
  ###### (6hrs+) LOOP LOOP End >> Getting All Combination outputs (Stepwise Regression Looping) for Selected Cluster ######
  
  

  ###### (6hrs+) LOOP > Export >> Write the List of Combinations to working Directory ######
  #write.csv (`Stepwise.Combination.Summary`, file = paste0("2b. Stepwise Combination Summary _15", Selected.Cluster,".csv"))

  
  
  ###### (6hrs+) LOOP > Identify the Best Combination for Each Cluster ######
  # Based on p value < 0.0001 and highest adjusted rsquare pick the model from the best combination
  # In the instant that there is more than one combination that results in the highest adjusted rsquared
  # -- pick the fist combination
  # -- This is due to the final model will have the same variables
  
  Stepwise.Combination.Summary <- subset (  Stepwise.Combination.Summary, 
                                            Combination.p.value < 0.0001)
  
  Final.Combination <- as.numeric (min(Stepwise.Combination.Summary$Combination[Stepwise.Combination.Summary$Combination.adj.r.squared==max(Stepwise.Combination.Summary$Combination.adj.r.squared)]))
  
  Final.Stepwise.Combination.by.Cluster <- rbind(Final.Stepwise.Combination.by.Cluster,
                                                 data.frame (Selected.Cluster, Final.Combination))
  
  rm(Stepwise.Combination.Summary)
  rm(Training.Set)
  rm(Testing.Set.Full.Range)
  rm(Testing.Set.Within.Range)
  rm(Final.Combination)
  rm(Selected.Cluster)
  
})


rm(Salesperson.Clustering)
names(Final.Stepwise.Combination.by.Cluster) <- c("Cluster", "Final Combination")

######## (6hrs+) LOOP END >> All Clusters - Finding the Best Combination of Inputs for Stepwise Regression (Run Once then skip) ########


###### Export >> Write the Best Combination of Inputs for Stepwise Regression to working Directory ######
write.csv (`Final.Stepwise.Combination.by.Cluster`, "2c. Final Stepwise Combination by Cluster_15.csv")








##### NEED TO UPDATE FROM HERE #####

##### CREATE OUTPUT ######
##### CREATE Graphs? ######
##### CREATE SUMMARY ######
##### Separate or Together? ######


#Model.Output.KPI.Refresh.All
#Model.Output.KPI.Refresh.Summary


######## <<< Manual>>> INTSTEAD OF LOOP - Use this output from the loop instead of running loop for further script if needed ########
Final.Stepwise.Combination.by.Cluster <- read.csv("2c. Final Stepwise Combination by Cluster_15.csv", check.names = FALSE)




######## <<Manual>> Select Cluster ########

# Selected.Cluster = "Drug Store Specialist Sales Route Type"
# Selected.Cluster = "General Trade Account Specialist Sales Route Type"
# Selected.Cluster = "Pre-Selling Public Market and Multi Channel Sales Route Type"
# Selected.Cluster = "Pre-Selling Sari Sari Stores Sales Route Type"
# Selected.Cluster = "Van Public Market and Multi Channel Sales Route Type"
#Selected.Cluster = "Van Sari Sari Stores Sales Route Type"

# Selected.Cluster = "HCF Product Group"
# Selected.Cluster = "MIX Product Group"
# Selected.Cluster = "PC Product Group"
# Selected.Cluster = "SPW Product Group"
gsv_data <- data.frame()
chart_data <- data.frame()

Salesperson.Clustering <- as.character(na.omit(as.character(unique(Final.Data.Stepwise.KPI$`Salesperson Clustering`))))
#Salesperson.Clustering <- Salesperson.Clustering[1:3]
for (Selected.Cluster in Salesperson.Clustering) {
  print(Selected.Cluster)
######## Selected Cluster Final Stepwise Regression Data Sets ########

Stepwise.Input.Cluster <-  subset ( Final.Data.Stepwise.KPI, `Salesperson Clustering` == Selected.Cluster)

# Only Complete Rows
Stepwise.Input.Cluster.OmitNA <- na.omit(Stepwise.Input.Cluster)

# Subset only Training Months - Measures Only
Training.Set <- subset (Stepwise.Input.Cluster.OmitNA, `Training / Testing / Removed Months` == "KPI Refresh - Training Months")
Training.Set <- subset (Training.Set, select = -c(`Salesperson Clustering` , `Month`, `Training / Testing / Removed Months`))

# Subset only Testing Months - Measures Only
Testing.Set.Full.Range <- subset (Stepwise.Input.Cluster.OmitNA, `Training / Testing / Removed Months` == "KPI Refresh - Testing Months")
Testing.Set.Full.Range <- subset (Testing.Set.Full.Range, select = -c(`Salesperson Clustering`, `Month`, `Training / Testing / Removed Months`))

# Subset rows in testing months where independent variables are within Range of tesing Set
Measures.Names <- names (subset (Training.Set, select = -c(`% Target - GSV%`)))
Testing.Set.Within.Range <- Testing.Set.Full.Range

for (Measure in Measures.Names) {
  max.range <-  max (`Training.Set`[which ( colnames (`Training.Set`) ==  Measure )])
  min.range <-  min (`Training.Set`[which ( colnames (`Training.Set`) ==  Measure )])
  `Testing.Set.Within.Range`[which ( colnames (`Testing.Set.Within.Range`) ==  Measure )] <- sapply (
    (`Testing.Set.Within.Range`[which ( colnames (`Testing.Set.Within.Range`) ==  Measure )]), function (x) {
      ifelse (x < min.range | x > max.range, NA, x)})
  rm(Measure)
  rm(min.range)
  rm(max.range)
}

rm(Measures.Names)

# Make Sure to change to numeric (same as training set) that it can be tested later
Testing.Set.Within.Range <- na.omit(Testing.Set.Within.Range)
Testing.Set.Within.Range <- data.frame(sapply(Testing.Set.Full.Range,as.numeric), check.names = FALSE)

rm(Stepwise.Input.Cluster)
rm(Stepwise.Input.Cluster.OmitNA)
rm(Testing.Set.Full.Range)


######## Final Regression Model  ########

library(MASS)

Final.Combination <-  subset ( Final.Stepwise.Combination.by.Cluster, `Cluster` == Selected.Cluster)
Final.Combination <- Final.Combination$`Final Combination`

Final.Stepwise.Cluster.Input <- subset ( Training.Set, select = c(List.of.Combination[[Final.Combination]])) 

Final.Stepwise.Cluster.Fit  <- lm( `% Target - GSV%` ~ . ,  data = `Final.Stepwise.Cluster.Input`)
Final.Stepwise.Cluster.Step <- stepAIC(Final.Stepwise.Cluster.Fit, direction="both",trace=0) 
# Final.Cluster.Step$anova 
# colnames(Final.Cluster.Step$model) 
Final.Stepwise.Cluster.Data <- Final.Stepwise.Cluster.Step$model
Final.Stepwise.Cluster.Model <-lm( `% Target - GSV%` ~., data=Final.Stepwise.Cluster.Data)
summary (Final.Stepwise.Cluster.Model)
rm(Final.Combination)

Final.Data.Test.and.Train <- rbind (Training.Set, Testing.Set.Within.Range)
Final.Data.Test.and.Train$Cluster <- Selected.Cluster
Final.Data.Test.and.Train$Country <- "Philippines"
chart_data <- rbind(chart_data,Final.Data.Test.and.Train)

######## RMSE for Final Regression Model ########
rm(Final.Stepwise.Cluster.Input)
rm(Final.Stepwise.Cluster.Fit)
rm(Final.Stepwise.Cluster.Step)
rm(Final.Stepwise.Cluster.Data)

Stepwise.Cluster.Test.Prediction <- predict (Final.Stepwise.Cluster.Model, 
                                             Testing.Set.Within.Range)
RMSE <- sqrt ( mean ( (Stepwise.Cluster.Test.Prediction - Testing.Set.Within.Range$`% Target - GSV%`)^2))
RME <- mean ( Stepwise.Cluster.Test.Prediction - Testing.Set.Within.Range$`% Target - GSV%`)

rm(Stepwise.Cluster.Test.Prediction)



######## TBC Validate Coeeficient / Rsquare / P Value Change ########

######## 1/ Summarising Regression and Interpreting Stepwise Coefficients ########
Final.Data.Test.and.Train <- rbind (Training.Set, Testing.Set.Within.Range)
rm (Training.Set)
rm (Testing.Set.Within.Range)

`Important KPIs / Measures for % Achievement of Target GSV` <- sapply ( variable.names(Final.Stepwise.Cluster.Model)[-1], function (x) {substr(x, 2, nchar(x)-1)})

`KPIs / Measures to Focus On`  <- sapply ( variable.names(Final.Stepwise.Cluster.Model)[-1], function (x) {
                                          ifelse ( x %in% "`Salesperson Target`", "Not Applicable",
                                                   ifelse ( (summary ( Final.Stepwise.Cluster.Model)$coefficients[x, c(4)] <= 0.001) & (summary (Final.Stepwise.Cluster.Model)$coefficients[x, c(1)] >= 0) , "Primary Recommendation",
                                                            ifelse ( (summary ( Final.Stepwise.Cluster.Model)$coefficients[x, c(4)] <= 0.05) & (summary (Final.Stepwise.Cluster.Model)$coefficients[x, c(1)] >= 0) , "Secondary Recommendation",
                                                                     ifelse ( (summary ( Final.Stepwise.Cluster.Model)$coefficients[x, c(4)] <= 0.001) & (summary (Final.Stepwise.Cluster.Model)$coefficients[x, c(1)] < 0) , "Check Quality of Measure **",
                                                                     "Not Recommended" ))))})
                                              
`% Achievement of Target GSV to 1 KPIs / Measures Input Unit`  <- sapply ( variable.names(Final.Stepwise.Cluster.Model)[-1], function (x) {round((summary(Final.Stepwise.Cluster.Model)$coefficients[x, c(1)]),4)})
`# KPIs / Measures Input Units to 1 % Achievement of Target GSV *`  <- sapply ( variable.names(Final.Stepwise.Cluster.Model)[-1], function (x) {round(1/(summary(Final.Stepwise.Cluster.Model)$coefficients[x, c(1)]),1)})
`Average Level of KPIs / Measures` <- sapply ( `Important KPIs / Measures for % Achievement of Target GSV`, function (x) {round(mean(Final.Data.Test.and.Train[,c(x)]),1)})

Interpreting.Coefficients <-  cbind ( data.frame (`Important KPIs / Measures for % Achievement of Target GSV`, check.names = FALSE),
                                      data.frame (`KPIs / Measures to Focus On`, check.names = FALSE),
                                      data.frame (`% Achievement of Target GSV to 1 KPIs / Measures Input Unit`, check.names = FALSE),
                                      data.frame (`# KPIs / Measures Input Units to 1 % Achievement of Target GSV *`, check.names = FALSE),
                                      data.frame (`Average Level of KPIs / Measures`, check.names = FALSE))

Interpreting.Coefficients <- rbind ( subset(Interpreting.Coefficients, `KPIs / Measures to Focus On` %in% "Primary Recommendation"),
                                     subset(Interpreting.Coefficients, `KPIs / Measures to Focus On` %in% "Secondary Recommendation"),
                                     subset(Interpreting.Coefficients, `KPIs / Measures to Focus On` %in% "Check Quality of Measure **"),
                                     subset(Interpreting.Coefficients, `KPIs / Measures to Focus On` %in% "Not Recommended"),
                                     subset(Interpreting.Coefficients, `KPIs / Measures to Focus On` %in% "Not Applicable"))
rownames(Interpreting.Coefficients) <- NULL
#View(Interpreting.Coefficients)  
Interpreting.Coefficients$`KPIs / Measures to Focus On` <- ifelse(Interpreting.Coefficients$`KPIs / Measures to Focus On`=="Check Quality of Measure **","Not Recommended **" ,as.character(Interpreting.Coefficients$`KPIs / Measures to Focus On`))

rm(`Important KPIs / Measures for % Achievement of Target GSV`)
rm(`KPIs / Measures to Focus On`)
rm(`% Achievement of Target GSV to 1 KPIs / Measures Input Unit`)
rm(`# KPIs / Measures Input Units to 1 % Achievement of Target GSV *`)
rm(`Average Level of KPIs / Measures`)

#gsv_data for UI
Interpreting.Coefficients$Cluster <- Selected.Cluster
Interpreting.Coefficients$Variability <- round(100*summary(Final.Stepwise.Cluster.Model)$adj.r.squared,1)
Interpreting.Coefficients$RMSE <- round(RMSE , 1)
Interpreting.Coefficients$Country <- "Philippines"
Interpreting.Coefficients$Train_start <- KPI.Refresh.Train.Start.Month
Interpreting.Coefficients$Train_end <- KPI.Refresh.Train.End.Month
Interpreting.Coefficients$Test_start <- KPI.Refresh.Test.Start.Month
Interpreting.Coefficients$Test_end <- KPI.Refresh.Test.End.Month
#Interpreting.Coefficients$`% Achievement of Target GSV to 1 KPIs / Measures Input Unit` <- NULL
gsv_data <- rbind(gsv_data,Interpreting.Coefficients)

}

chart_data <- chart_data[,c(34,1:33,35)]
Inputs <- data.frame(Cluster=Salesperson.Clustering,Country="Philippines")
Chart_var <- data.frame(Varnames=colnames(chart_data))
Ques <- data.frame(Select=c("Which KPI and Measures should we incentivise to increase Distributor Sales Representative (DSR) % Achievement of Target GSV?",
                            "What is the Maximum / Minimum Level for the recommended KPIs and Measures to incentivise?"))



wb <- openxlsx::createWorkbook(creator = "")

openxlsx::addWorksheet(wb, sheetName = "Inputs")
openxlsx::addWorksheet(wb, sheetName = "GSV_Data")
openxlsx::addWorksheet(wb, sheetName = "Chart_Data")
openxlsx::addWorksheet(wb, sheetName = "Chart_var")
openxlsx::addWorksheet(wb, sheetName = "Ques")
openxlsx::writeData(wb, sheet = 1, x = Inputs)
openxlsx::writeData(wb, sheet = 2, x = gsv_data)
openxlsx::writeData(wb, sheet = 3, x = chart_data)
openxlsx::writeData(wb, sheet = 4, x = Chart_var)
openxlsx::writeData(wb, sheet = 5, x = Ques)
Sys.setenv("R_ZIPCMD" ="C:/Rtools/bin/zip.exe")
openxlsx::saveWorkbook(wb,paste0("Final_data_15_",gsub("[:punct:]","",Sys.Date()),".xlsx"),overwrite = T)



