# Load required packages (if you dont have them use install.packag --------

## Load required packages (if you dont have them use install.packages("nameofpackage") first 
library(haven)
library(tidyverse)
library(openxlsx)
library(readxl)
library(labelled)
library(rlang)
library(codebook)
library(conflicted)
# Import data set and create codebook -------------------------------------

#change directory to where your files are
setwd("C:\\Users\\idrissa.dabo\\OneDrive - World Food Programme\\Documents\\Custom Office Templates\\Tchad")
#import SPSS data set
data  <- read_sav("TCD_201910_ENSAM_external.sav")
#change values to value labels 
conflict_prefer("to_factor", "labelled")
data <- to_factor(data)
# create codebook so you can find your variables and variable labels - make sure to turn off
#codebook_browser(data)


# Select and rename the key variables and values used for the analysis --------

#standardize names of key variables using WFP VAM's Assesment Codebook
data <- data  %>% select(ADMIN1Name = region, #first adminstrative division
                                            ADMIN2Name = Strate, #second administrative division
                                            HDDScore = HDDS, #Household Dietary Diversity Score
                                            FCSCat = Gpe_SCA_Calcule_CH, #Food Consumption Groups from the Food Consumption Score 21/35 - normal threshold
                                            HHScore = Echelle_faim, #Household Hunger Score
                                            rCSIScore = rCSI, #reduced coping strategies
                                            LhHCSCat = Strategie_moyen_existence, #livelihood coping strategies
                                            #no weights
                                            #factor contributifs dessus - cant find many of the factors
                                            choc_subi, 
                                            Sup_detruite_ennemis_culture, #not sure this one really matches
                                            Indice_richesse_Final, 
                                            source_eau_boisson,
                                            nombre_repas_enfant,    
                                            energie_cuisson_aliment)


# # standardize naming of categorical values -----------------------------

# launch codebook agian so you can find see value labels - make sure to turn off
#codebook_browser(data)

#standardize naming of categorical values and also convert numeric values to Cadre Harmonise thresholds 
data <- data %>%  mutate(FCSCat = case_when(
  FCSCat == "Pauvre" ~ "Poor", 
  FCSCat == "Limite" ~ "Borderline",       
  FCSCat == "Acceptable" ~ "Acceptable"
),
LhHCSCat = case_when(
  LhHCSCat == "Aucune strategie" ~ "NoStrategies", # Put accents to strat?gies so that the code works properly
  LhHCSCat == "Strategie de stress" ~ "StressStrategies",
  LhHCSCat == "Strategie de crise" ~ "CrisisStrategies",
  LhHCSCat == "Strategie durgence" ~ "EmergencyStrategies"
),
CH_HDDS = case_when(
  HDDScore >= 5 ~ "Phase1", 
  HDDScore == 4 ~ "Phase2",       
  HDDScore == 3 ~ "Phase3",
  HDDScore == 2 ~ "Phase4",
  HDDScore < 2 ~ "Phase5"
),
CH_rCSI = case_when(
  rCSIScore <= 3 ~ "Phase1", 
  rCSIScore >= 4 & rCSIScore <= 18 ~ "Phase2",       
  rCSIScore >= 19 ~ "Phase3"    
),
CH_HHS =
  case_when(
    HHScore == 0 ~ "Phase1",
    HHScore == 1 ~ "Phase2",
    HHScore == 2 | HHScore == 3 ~ "Phase3",
    HHScore == 4 ~ "Phase4",  
    HHScore >= 5 ~ "Phase5"
  ))

# create tables of the proportion by administrative -----------------------

##create tables of the proportion by administrative areas and then apply the CH indicator specific and 20% rule to each indicator
#Household Dietarty Diversity Score
HDDStable <- data %>% 
  group_by(ADMIN1Name, ADMIN2Name) %>%
  count(CH_HDDS) %>%
  drop_na() %>%
  mutate(n = 100 * n / sum(n)) %>%
  ungroup() %>%
  pivot_wider(names_from = CH_HDDS,
              values_from = n,
              values_fill = list(n=0)) %>% 
  mutate_if(is.numeric, round, 1)
#Apply the 20% rule (if it is 20% in that phase or the sum of higher phases equals 20%) 
CH_HDDStable <- HDDStable %>% mutate(indicator = "HDDS", 
                                     phase2345 = `Phase2` + `Phase3` + `Phase4` + `Phase5`, #this variable will be used to see if phase 2 and higher phases equals 20% or more
                                     phase345 = `Phase3` + `Phase4` + `Phase5`, #this variable will be used to see if phase 3 and higher phases equal 20% or more
                                     phase45 = `Phase4` + `Phase5`, #this variable will be used to see if phase 3 and higher phases equal 20% or more
                                     finalphase = case_when(
                                       `Phase5` >= 20 ~ 5, #if 20% or more is in phase 5 then assign phase 5
                                       `Phase4` >= 20 | phase45 >= 20 ~ 4, #if 20% or more is in phase 4 or the sum of phase4 and 5 is more than 20% then assign phase 4
                                       `Phase3` >= 20 | phase345 >= 20 ~ 3, #if 20% or more is in phase 3 or the sum of phase3, 4 and 5 is more than 20% then assign phase 3
                                       `Phase2` >= 20 | phase2345 >= 20 ~ 2, #if 20% or more is in phase 2 or the sum of phase 2, 3, 4 and 5 is more than 20% then assign phase 2
                                       TRUE ~ 1)) %>% #otherwise assign phase 1
  select(indicator, ADMIN1Name, ADMIN2Name, HDDS_Phase1 = Phase1, HDDS_Phase2 = Phase2, HDDS_Phase3 = Phase3, HDDS_Phase4 = Phase4, HDDS_Phase5 = Phase5, HDDS_finalphase = finalphase) #select only relevant variables, rename them with indicator name and order in proper sequence

#Food Consumption Groups
FCGtable <- data %>% 
  group_by(ADMIN1Name, ADMIN2Name) %>%
  count(FCSCat) %>%
  drop_na() %>%
  mutate(n = 100 * n / sum(n)) %>%
  ungroup() %>%
  pivot_wider(names_from = FCSCat,
              values_from = n,
              values_fill = list(n=0)) %>% 
  mutate_if(is.numeric, round, 1)
#Apply the Cadre Harmonise rules for phasing the Food Consumption Groups 
CH_FCGtable <-  FCGtable %>%  mutate(indicator = "FCG", PoorBorderline = Poor + Borderline, finalphase = case_when(
  Poor < 5 ~ 1,  #if less than 5% are in the poor food group then phase 1
  Poor >= 20 ~ 4, #if 20% or more are in the poor food group then phase 4
  between(Poor,5,10) ~ 2, #if % of people are between 5 and 10%  then phase2
  between(Poor,10,20) & PoorBorderline < 30 ~ 2, #if % of people in poor food group are between 20 and 20% and the % of people who are in poor and borderline is less than 30 % then phase2
  between(Poor,10,20) & PoorBorderline >= 30 ~ 3)) %>% #if % of people in poor food group are between 20 and 20% and the % of people who are in poor and borderline is less than 30 % then phase2
  select(indicator, ADMIN1Name, ADMIN2Name, FCG_Poor = Poor, FCG_Borderline = Borderline, FCG_Acceptable = Acceptable, FCG_finalphase = finalphase) #select only relevant variables and order in proper sequence

#Household Hunger Score
HHStable <- data %>% 
  group_by(ADMIN1Name, ADMIN2Name) %>%
  count(CH_HHS) %>%
  drop_na() %>%
  mutate(n = 100 * n / sum(n)) %>%
  ungroup() %>%
  pivot_wider(names_from = CH_HHS,
              values_from = n,
              values_fill = list(n=0)) %>% 
  mutate_if(is.numeric, round, 1)
#Apply the 20% rule (if it is 20% in that phase or the sum of higher phases equals 20%) 
CH_HHStable <- HHStable %>% mutate(indicator = "HHS", phase2345 = `Phase2` + `Phase3` + `Phase4` + `Phase5`,
                                   phase345 = `Phase3` + `Phase4` + `Phase5`,
                                   phase45 = `Phase4` + `Phase5`,
                                   finalphase = case_when(
                                     Phase5 >= 20 ~ 5,
                                     Phase4 >= 20 | phase45 >= 20 ~ 4,
                                     Phase3 >= 20 | phase345 >= 20 ~ 3,
                                     Phase2 >= 20 | phase2345 >= 20 ~ 2,
                                     TRUE ~ 1)) %>% 
  select(indicator, ADMIN1Name, ADMIN2Name, HHS_Phase1 = Phase1, HHS_Phase2 = Phase2, HHS_Phase3 = Phase3, HHS_Phase4 = Phase4, HHS_Phase5 = Phase5, HHS_finalphase = finalphase)

#reduced consumption score
rCSItable <- data %>% 
  group_by(ADMIN1Name, ADMIN2Name) %>%
  count(CH_rCSI) %>%
  drop_na() %>%
  mutate(n = 100 * n / sum(n)) %>%
  ungroup() %>%
  pivot_wider(names_from = CH_rCSI,
              values_from = n,
              values_fill = list(n=0)) %>% 
  mutate_if(is.numeric, round, 1)
#Apply the 20% rule (if it is 20% in that phase or the sum of higher phases equals 20%) 
CH_rCSItable <- rCSItable %>% mutate(indicator = "rCSI", 
                                     rcsi23 = Phase2 + Phase3,
                                     finalphase =
                                       case_when(
                                         Phase3 >= 20 ~ 3, 
                                         Phase2 >= 20 | rcsi23 >= 20 ~ 2,
                                         TRUE ~ 1
                                       )) %>% select(indicator, ADMIN1Name, ADMIN2Name, rCSI_Phase1 = Phase1, rCSI_Phase2 = Phase2, rCSI_Phase3 =Phase3, rCSI_finalphase = finalphase)


#Livelihood Coping Strategies
LhHCSCattable <- data %>% 
  group_by(ADMIN1Name, ADMIN2Name) %>%
  count(LhHCSCat) %>%
  drop_na() %>%
  mutate(n = 100 * n / sum(n)) %>%
  ungroup() %>%
  pivot_wider(names_from = LhHCSCat,
              values_from = n,
              values_fill = list(n=0)) %>% 
  mutate_if(is.numeric, round, 1)
#Apply the Cadre Harmonise rules for phasing the Livelihood Coping Strategies 
CH_LhHCSCattable <- LhHCSCattable %>% mutate(indicator = "LhHCSCat", stresscrisisemergency = StressStrategies + CrisisStrategies + EmergencyStrategies,
                                             crisisemergency = CrisisStrategies + EmergencyStrategies,
                                             finalphase = case_when(
                                               EmergencyStrategies >= 20 ~ 4,
                                               crisisemergency >= 20 & EmergencyStrategies < 20 ~ 3,  
                                               NoStrategies < 80 & crisisemergency < 20 ~ 2,
                                               NoStrategies >= 80 ~ 1
                                             )) %>% select(indicator, ADMIN1Name, ADMIN2Name, LhHCSCat_NoStrategies = NoStrategies, LhHCSCat_StressStrategies = StressStrategies, LhHCSCat_CrisisStategies = CrisisStrategies, LhHCSCat_EmergencyStrategies = EmergencyStrategies, LhHCSCat_finalphase = finalphase)

# Add contributing factors variables --------------------------------------

##Add contributing factors variables (different from the Food Security direct evidence above, these variables will depend country to country)
##so that the contributing factors can be imported into the proper category, the final variable names should be given a prefix (e.g. 01_, 02_)
##"Hazards & Vulnerability" = 01 - 10
##"Availibility" = 11 - 25
##"Accessibility" = 26 - 40
##"Utilization including access to clean water" = 41 - 55
##"Stability" = 56 - 70

#Create a table of the proportion of people who "experienced a shock in the last six months"
choc_subi_table <- data %>% 
  group_by(ADMIN1Name, ADMIN2Name) %>%  mutate(indicator = "Pendant les six dernier mois, le menage a-t-il subi un choc?") %>%
  count(choc_subi) %>%
  drop_na() %>%
  mutate(n = 100 * n / sum(n)) %>%
  ungroup() %>%
  pivot_wider(names_from = choc_subi,
              values_from = n,
              values_fill = list(n=0)) %>% 
  mutate_if(is.numeric, round, 1) %>% 
  select(ADMIN1Name, ADMIN2Name, `01_choc_subi_Non` = Non, `01_choc_subi_Oui` = Oui)

Sup_detruite_ennemis_culture_table <- data %>% drop_na(Sup_detruite_ennemis_culture) %>%
  group_by(ADMIN1Name, ADMIN2Name) %>% mutate(indicator = "Quelle est la principale source d'eau de boisson de votre mÃ©nage ?") %>%
  summarise(moyenne_Sup_detruite_ennemis_culture = mean(Sup_detruite_ennemis_culture)) %>% ungroup() %>%
  select(ADMIN1Name, ADMIN2Name, `02_moyenne_Sup_detruite_ennemis_culture` = moyenne_Sup_detruite_ennemis_culture)

Indice_richesse_Final_table <- data %>% 
  group_by(ADMIN1Name, ADMIN2Name) %>%  mutate(indicator = "Quelle est la principale source d'eau de boisson de votre mÃ©nage ?") %>%
  count(Indice_richesse_Final) %>%
  drop_na() %>%
  mutate(n = 100 * n / sum(n)) %>%
  ungroup() %>%
  pivot_wider(names_from = Indice_richesse_Final,
              values_from = n,
              values_fill = list(n=0)) %>% 
  mutate_if(is.numeric, round, 1) %>% 
  select(ADMIN1Name, ADMIN2Name, `26_Indice_richesse_Final_Quintile de richesse tres pauvre` = `Quintile de richesse tres pauvre`,
         `26_Indice_richesse_Final_Quintile de richesse pauvre` = `Quintile de richesse pauvre`, `26_Indice_richesse_Final_Quintile de richesse moyen` = `Quintile de richesse moyen`,
         `26_Indice_richesse_Final_Quintile de richesse nanti` = `Quintile de richesse nanti`)


source_eau_boisson_table <- data %>% 
  group_by(ADMIN1Name, ADMIN2Name) %>%  mutate(indicator = "Quelle est la principale source d'eau de boisson de votre mÃ©nage ?") %>%
  count(source_eau_boisson) %>%
  drop_na() %>%
  mutate(n = 100 * n / sum(n)) %>%
  ungroup() %>%
  pivot_wider(names_from = source_eau_boisson,
              values_from = n,
              values_fill = list(n=0)) %>% 
  mutate_if(is.numeric, round, 1) %>% 
  select(ADMIN1Name, ADMIN2Name, `41_source_eau_boisson_Robineteaucourante` = `Robinet eau courante`, `41_source_eau_boisson_Forage/pompe` = "Forage/pompe",
         `41_source_eau_boisson_Eaudesurface` = `Eau de surface`, `41_source_eau_boisson_Puitsamélioré` = `Puits amélioré`, `41_source_eau_boisson_Puitstraditionnel` = `Puits traditionnel`)


nombre_repas_enfant_table <- data %>% drop_na(nombre_repas_enfant) %>%
  group_by(ADMIN1Name, ADMIN2Name) %>% mutate(indicator = "Quelle est la principale source d'eau de boisson de votre mÃ©nage ?") %>%
  summarise(moyenne_nombre_repas_enfant = mean(nombre_repas_enfant)) %>% ungroup() %>%
  select(ADMIN1Name, ADMIN2Name, `42_moyenne_nombre_repas_enfant` = moyenne_nombre_repas_enfant)
  
energie_cuisson_aliment_table <- data %>% 
  group_by(ADMIN1Name, ADMIN2Name) %>%  mutate(indicator = "Quelle est la principale source d'eau de boisson de votre mÃ©nage ?") %>%
  count(energie_cuisson_aliment) %>%
  drop_na() %>%
  mutate(n = 100 * n / sum(n)) %>%
  ungroup() %>%
  pivot_wider(names_from = energie_cuisson_aliment,
              values_from = n,
              values_fill = list(n=0)) %>% 
  mutate_if(is.numeric, round, 1) %>% 
  select(ADMIN1Name, ADMIN2Name, `43_energie_cuisson_aliment_Bois` = `Bois`, `43_energie_cuisson_aliment_Charbon de bois` = `Charbon de bois`,
         `43_energie_cuisson_aliment_Gaz` = `Gaz`, `43_energie_cuisson_aliment_Electricité` = `Electricité`, `43_energie_cuisson_aliment_Déchets des animaux` = `Déchets des animaux`, `43_energie_cuisson_aliment_Autres` = `Autres`)

# Merge key variables -----------------------------------------------------

### Merge key variables from Direct Evidence and Contributing factor tables together - problem is that some factors contributif have missing values
matrice_intermediaire <- bind_cols(
  select(CH_FCGtable,-"indicator"),# select all variables except indicator 
  select(CH_HDDStable,-c("indicator","ADMIN1Name","ADMIN2Name")),# select all variables except indicator, ADMIN1Name and ADMIN2Name because the latter two are already selected from the table CH_FCGtable
  select(CH_HHStable,-c("indicator","ADMIN1Name","ADMIN2Name")),
  select(CH_LhHCSCattable,-c("indicator","ADMIN1Name","ADMIN2Name")),
  select(CH_rCSItable,-c("indicator","ADMIN1Name","ADMIN2Name")),
  select(choc_subi_table,-c("ADMIN1Name","ADMIN2Name")),
  select(Sup_detruite_ennemis_culture_table,-c("ADMIN1Name","ADMIN2Name")),
  select(Indice_richesse_Final_table,-c("ADMIN1Name","ADMIN2Name")),
  select(source_eau_boisson_table,-c("ADMIN1Name","ADMIN2Name")),
  select(nombre_repas_enfant_table,-c("ADMIN1Name","ADMIN2Name")),
  select(energie_cuisson_aliment_table,-c("ADMIN1Name","ADMIN2Name")),
  )
  
#add the one contributing factor that has missing values
#matrice_intermediaire <- full_join(matrice_intermediaire,secheresse_derniere_camp_table,by=c("ADMIN1Name","ADMIN2Name")) %>% 
# replace(., is.na(.), " ")

# create a blank space for other variables --------------------------------

###create a blank space for other variables that will be used in the excel analyis but do not come from the survey data above


# create variables 
Z1_DPME_C  # DPME_zone1(courant)
Z1_DPME_pop_C  #% Pop DPME_Zone1(courant)
Z1_DS_C  # DS_Zone1(courant)
Z1_Pop_DS_C  # %Pop DS_Zone1(courant)
Z1_DPME_Pr  # DPME_Zone1(projetee)
Z1_Pop_DPME_Pr # %Pop DPME_Zone1(projetee)
Z1_DS_Pr  # DS_Zone1(projetee)
Z1_pop_DS_Pr  # %Pop DS_Zone1(projetee)

Z2_DPME_C  # DPME_zone2(courant)
Z2_DPME_pop_C  # %Pop DPME_Zone2(courant)
Z2_DS_C  # DS_Zone2(courant)
Z2_Pop_DS_C # %Pop DS_Zone2(courant)
Z2_DPME_Pr  # DPME_Zone2(projetee)
Z2_Pop_DPME_Pr  # %Pop DPME_Zone2(projetee)
Z2_DS_Pr  # DS_Zone2(projetee)
Z2_pop_DS_Pr  # %Pop DS_Zone2(projetee)

Z3_DPME_C  # DPME_zone3(courant)
Z3_DPME_pop_C  #% Pop DPME_Zone3(courant)
Z3_DS_C  # DS_Zone3(courant)
Z3_Pop_DS_C  # %Pop DS_Zone3(courant)
Z3_DPME_Pr  # DPME_Zone3(projetee)
Z3_Pop_DPME_Pr  # %Pop DPME_Zone3(projetee)
Z3_DS_Pr  # DS_Zone3(projetee)
Z3_pop_DS_Pr  # %Pop DS_Zone3(projetee)

Z4_DPME_C  # DPME_zone4(courant)
Z4_DPME_pop_C  #% Pop DPME_Zone4(courant)
Z4_DS_C  # DS_Zone4(courant)
Z4_Pop_DS_C  # %Pop DS_Zone4(courant)
Z4_DPME_Pr   # DPME_Zone4(projetee)
Z4_Pop_DPME_Pr  # %Pop DPME_Zone4(projetee)
Z4_DS_Pr  # DS_Zone1(projetee)
Z4_pop_DS_Pr  # %Pop DS_Zone1(projetee)

Proxy_cal  # Proxy calorique
MAG_pt  # MAG-P/T
MAG_Pharv  # MAG-Midian
MAG_Soud  # MAG-Midiane soudure
IMC  # IMC
MUAC  # MAG-MUAC 
TBM  # TBM
TMM5  # TMM5
Population  # Population
Geocode  # Geocode


# Add the other variables to table containing direct evidence and contributing factors  --------

matrice_intermediaire <- matrice_intermediaire %>% 
  mutate(Z1_DPME_C=" ",Z1_DPME_pop_C=" ",Z1_DS_C=" ",Z1_Pop_DS_C=" ",  
         Z1_DPME_Pr=" ",Z1_Pop_DPME_Pr=" ",Z1_DS_Pr=" ",Z1_pop_DS_Pr=" ",
         Z2_DPME_C=" ",Z2_DPME_pop_C=" ",Z2_DS_C=" ",Z2_Pop_DS_C=" ",
         Z2_DPME_Pr=" ",Z2_Pop_DPME_Pr=" ",Z2_DS_Pr=" ",Z2_pop_DS_Pr=" ",
         Z3_DPME_C=" ",Z3_DPME_pop_C=" ",Z3_DS_C=" ",Z3_Pop_DS_C=" ",
         Z3_DPME_Pr=" ",Z3_Pop_DPME_Pr=" ",Z3_DS_Pr=" ",Z3_pop_DS_Pr=" ",
         Z4_DPME_C=" ",Z4_DPME_pop_C=" ",Z4_DS_C=" ",Z4_Pop_DS_C=" ",
         Z4_DPME_Pr=" ",Z4_Pop_DPME_Pr=" ",Z4_DS_Pr=" ",Z4_pop_DS_Pr=" ",
         Proxy_cal=" ", MAG_pt=" ", MAG_Pharv=" ", MAG_Soud=" ",
         IMC=" ", MUAC=" ", TBM=" ", TMM5=" ",Population=" ",Geocode=" ")


#re-orders the table
matrice_intermediaire <- matrice_intermediaire %>% 
  select(ADMIN1Name,ADMIN2Name,Population,Geocode,everything())


# saving final data as excel sheet ----------------------------------------

direct <- createStyle(fgFill = "#4F81BD", halign = "left", textDecoration = "Bold",
                      border = "Bottom", fontColour = "white")
contributifs <- createStyle(fgFill = "#FFC7CE", halign = "left", textDecoration = "Bold",
                            border = "Bottom", fontColour = "black")
hea <- createStyle(fgFill = "#C6EFCE", halign = "left", textDecoration = "Bold",
                   border = "Bottom", fontColour = "black")

nutrition <- createStyle(fgFill = "yellow", halign = "left", textDecoration = "Bold",
                         border = "Bottom", fontColour = "black")

mortalite <- createStyle(fgFill = "orange", halign = "left", textDecoration = "Bold",
                         border = "Bottom", fontColour = "black")

proxyvar <- createStyle(fgFill = "lightgreen", halign = "left", textDecoration = "Bold",
                     border = "Bottom", fontColour = "black")

col1stDirectVariable <- which(colnames(matrice_intermediaire)=="FCG_Poor")
colLastDirectVariable <- which(colnames(matrice_intermediaire)=="rCSI_finalphase")
col1stHEAVariable <- which(colnames(matrice_intermediaire)=="Z1_DPME_C")
colLastHEAVarible <- which(colnames(matrice_intermediaire)=="Z4_pop_DS_Pr")
col1stNutritionVariable <- which(colnames(matrice_intermediaire)=="MAG_pt")
colLastNutritionVariable <- which(colnames(matrice_intermediaire)=="MUAC")
col1stMortaliteVariable <- which(colnames(matrice_intermediaire)=="TBM")




Matrice_intermediaire <- createWorkbook()
addWorksheet(Matrice_intermediaire, "Matrice intermediaire")
writeData(Matrice_intermediaire,sheet = 1,matrice_intermediaire,startRow = 1,startCol = 1,)
addStyle(Matrice_intermediaire,1,rows = 1,cols = col1stDirectVariable:colLastDirectVariable
         ,style = direct,gridExpand = TRUE,)
addStyle(Matrice_intermediaire,1,rows = 1,cols =(colLastDirectVariable+1) :(col1stHEAVariable-1),
         style = contributifs,gridExpand = TRUE,)
addStyle(Matrice_intermediaire,1,rows = 1,cols =col1stHEAVariable :colLastHEAVarible,
         style = hea,gridExpand = TRUE,)
addStyle(Matrice_intermediaire,1,rows = 1,cols =col1stNutritionVariable :colLastNutritionVariable,
         style = nutrition,gridExpand = TRUE,)
addStyle(Matrice_intermediaire,1,rows = 1,cols =col1stNutritionVariable-1 ,
         style = proxyvar,gridExpand = TRUE,)
addStyle(Matrice_intermediaire,1,rows = 1,cols = col1stMortaliteVariable:ncol(matrice_intermediaire),
         style = mortalite,gridExpand = TRUE,)

saveWorkbook(Matrice_intermediaire,file ="Matrice_intermediaire.xlsx",overwrite = TRUE)
 openXL(Matrice_intermediaire)
