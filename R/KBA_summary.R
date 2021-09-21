form_conversion <- function(KBAforms, includeQuestions, includeReviewDetails){


 Packages
library(tidyverse)
library(magrittr)
library(WordR)
library(flextable)
library(officer)
library(openxlsx)
library(lubridate)
library(janitor)
library(stringr)

# Options
options(scipen = 999)

# Data
      # Species list
speciesList <- read.xlsx('joint_files/Ref_Species.xlsx', sheet=2) %>%
  mutate(IUCN_AssessmentDate = convertToDate(IUCN_AssessmentDate), COSEWIC_DATE = convertToDate(COSEWIC_DATE), G_RANK_REVIEW_DATE = convertToDate(G_RANK_REVIEW_DATE), N_RANK_REVIEW_DATE = convertToDate(N_RANK_REVIEW_DATE)) %>%
  select(NATIONAL_SCIENTIFIC_NAME, Endemism, IUCN_CD, IUCN_AssessmentDate, COSEWIC_STATUS, COSEWIC_DATE, ROUNDED_G_RANK, G_RANK_REVIEW_DATE, ROUNDED_N_RANK, N_RANK_REVIEW_DATE)
## Google Drive: https://docs.google.com/spreadsheets/d/1R2ILLvyGMqRL8S9pfZdYIeBKXlyzckKQ/edit?usp=sharing&ouid=104844399648613391324&rtpof=true&sd=true

      # Criteria definitions
criteria_definitions <- read.xlsx("joint_files/KBACriteria_Definitions.xlsx")
## Google Drive: https://docs.google.com/spreadsheets/d/1c-2sbnvOfp3hjw5UqVYKC64QmqW205B0/edit?usp=sharing&ouid=104844399648613391324&rtpof=true&sd=true


#### Prepare the Summary(ies) ####

for(step in 1:length(KBAforms)){
  
  # Load KBA Canada Proposal Form
        # Visible sheets
  home <- read.xlsx(KBAforms[step], sheet = "HOME")
  proposer <- read.xlsx(KBAforms[step], sheet = "1. PROPOSER")
  site <- read.xlsx(KBAforms[step], sheet = "2. SITE")
  species <- read.xlsx(KBAforms[step], sheet = "3. SPECIES")
  ecosystems <- read.xlsx(KBAforms[step], sheet = "4. ECOSYSTEMS & C")
  threats <- read.xlsx(KBAforms[step], sheet = "5. THREATS")
  review <- read.xlsx(KBAforms[step], sheet = "6. REVIEW")
  citations <- read.xlsx(KBAforms[step], sheet = "7. CITATIONS")
  check <- read.xlsx(KBAforms[step], sheet = "8. CHECK")
  
        # Invisible sheets
  checkboxes <- read.xlsx(KBAforms[step], sheet = "checkboxes")
  resultsSpecies <- read.xlsx(KBAforms[step], sheet = "results_species")
  resultsEcosystems <- read.xlsx(KBAforms[step], sheet = "results_ecosystems")

  # Get form version number
  formVersion <- home[1,1] %>% substr(., start=9, stop=nchar(.)) %>% as.numeric()
  
  # Format the sheets
        # 1. PROPOSER
  proposer %<>%
    .[ ,2:3] %>%
    rename(Field = X2, Entry = X3) %>%
    filter(!is.na(Field))
  
        # 2. SITE
  site %<>%
    .[ , 2:4] %>%
    rename(Field = X2) %>%
    filter(!Field == "Ongoing                                                                                           Needed                                                  ")
  
  actions <- checkboxes %>%
    .[,5:7]
  colnames(actions) <- actions[1,]
  actions %<>%
    .[2:nrow(.),]
  
        # 3. SPECIES
  colnames(species) <- species[1,]
  species %<>%
    .[2:nrow(.),] %>%
    filter(!is.na(`Common name`)) %>%
    mutate(`Common name` = trimws(`Common name`),
           `Scientific name` = trimws(`Scientific name`))
  
        # 4. ECOSYSTEMS & C
  ecosystems %<>%
    pull(X2) %>%
    unique() %>%
    .[which(!. == "Criteria met")]
  if(!length(ecosystems) == 0){stop("Ecosystem KBAs not yet supported. Please contact Chloé and provide her with the error message.")}
        
        # 5. THREATS
              # Verify whether "No Threats" checkbox is checked
  noThreats <- checkboxes[2,9] %>% as.logical()
  
              # If there are threats, get that information
  if(!noThreats){
    colnames(threats) <- threats[3,]
    threats %<>% .[4:nrow(.),]
    colnames(threats)[ncol(threats)] <- "Notes"
  }  
     
        # 6. REVIEW
  review %<>%
    drop_na(X2) %>%
    fill(`INSTRUCTIONS:`)
  
  technicalReview <- review %>%
    filter(`INSTRUCTIONS:` == 1) %>%
    select(-`INSTRUCTIONS:`)
  colnames(technicalReview) <- technicalReview[2,]
  if(nrow(technicalReview) > 2){
    technicalReview %<>% .[3:nrow(.),]
  }else{
    technicalReview[3,] <- c("No reviewers listed", "", "", "")
    technicalReview %<>% .[3:nrow(.),]
  }
  
  generalReview <- review %>%
    filter(`INSTRUCTIONS:` == 2) %>%
    select(-c(`INSTRUCTIONS:`, X5))
  colnames(generalReview) <- generalReview[2,]
  if(nrow(generalReview) > 2){
    generalReview %<>% .[3:nrow(.),]
  }else{
    generalReview[3,] <- c("No reviewers listed", "", "")
    generalReview %<>% .[3:nrow(.),]
  }
   
        # 7. CITATIONS
  colnames(citations) <- citations[2,]
  citations %<>%
    .[3:nrow(.), 1:4] %>%
    filter(!is.na(`Short Citation`))
  
        # 8. CHECK
              # Column names
  colnames(check) <- c("Check", "Item")
  
              # Get checkbox results
  check_checkboxes <- checkboxes %>%
    .[2:nrow(.),] %>%
    select("8..Checks") %>%
    drop_na()
  if(formVersion == 1.1){check_checkboxes %<>% .[c(1:5,7:nrow(.)),]} # Cell N8 is obsolete in v1.1 of the Proposal Form (it doens't link to any actual checkbox)
  
              # Verify that there are as many checkbox results as there are checkboxes
  if(!(nrow(check) == length(check_checkboxes))){stop("Inconsistencies between the 8. CHECKS tab and checkbox results. This error originates from the Excel formulas themselves. Please contact Chloé and provide her with the error message.")}
  
              # Add checkbox results to the 8. CHECK tab
  check %<>%
    select(-Check) %>%
    mutate(Check = check_checkboxes)
  rm(check_checkboxes)
  
  # Prepare variables
        # 1. KBA Name
  nationalName <- site$GENERAL[which(site$Field == "National name")]
  
        # 2. Location
              # Jurisdiction
  juris <- site$GENERAL[which(site$Field == "Province or Territory")]
  
              # Latitude and Longitude
  lat <- site$GENERAL[which(site$Field == "Latitude (dd.dddd)")] %>%
    as.numeric(.) %>%
    round(., 3)
  lat <- ifelse(is.na(lat), "coordinates unspecified", lat)
  
  lon <- site$GENERAL[which(site$Field == "Longitude (dd.dddd)")] %>%
    as.numeric(.) %>%
    round(., 3)
  lon <- ifelse(is.na(lon), "", paste0("/", lon))
  
        # 3. KBA Scope
  scope <- ifelse(grepl("g", home[13,4], fixed=T) & grepl("n", home[13,4], fixed=T),
                  "Global and National",
                  ifelse(grepl("g", home[13,4], fixed=T),
                         "Global",
                         "National"))
  
        # 4. Proposal Development Lead
  proposalLead <- proposer$Entry[which(proposer$Field == "Name")]
  
        # 7. Site Description
  siteDescription <- site$GENERAL[which(site$Field == "Site description")]
  
        # 8. Assessment Details - KBA Trigger Species
  includeGlobalTriggers <- ifelse(scope %in% c("Global and National", "Global"), "GLOBAL", "")
  includeNationalTriggers <- ifelse(scope %in% c("Global and National", "National"), "NATIONAL", "")
  
        # 10. Delineation Rationale
  delineationRationale <- site$GENERAL[which(site$Field == "Delineation rationale")]
  
        # 12. General Review
  noFeedback <- review$X3[which(review$X2 == "Provide information about any organizations you contacted and that did not provide feedback.")]
  noFeedback <- ifelse(is.na(noFeedback), "None", noFeedback)
  
        # 13. Additional Site Information
  nominationRationale <- site$GENERAL[which(site$Field == "Rationale for nomination")]
  additionalBiodiversity <- site$GENERAL[which(site$Field == "Additional biodiversity")]
  percentProtected <- site$GENERAL[which(site$Field == "Percent protected")]
  customaryJurisdiction <- site$GENERAL[which(site$Field == "Customary jurisdiction")]
  
  # Prepare flextables
        # Criteria information
              # Get data
  criteriaMet <- home$X4[which(home$X3 == "Criteria met")]
  
              # Check that at least one criterion is met
  if(is.na(criteriaMet)){
    stop("No KBA Criteria met. Please revise your form and ensure that at least one criterion is met. If you believe that a KBA criterion should be met based on the information you provided in the form, contact Chloé and provide her with the error message.")
  }
  
              # Criteria definitions
  criteriaInfo <- data.frame(CriteriaFull = strsplit(criteriaMet, "; ")[[1]]) %>%
    mutate(Scope = ifelse(grepl("g", CriteriaFull, fixed=T), "Global", "National")) %>%
    mutate(Criteria = sapply(CriteriaFull, function(x) substr(x, start=2, stop=nchar(x)))) %>%
    arrange(Scope, Criteria) %>%
    mutate(Definition = sapply(1:nrow(.), function(x) criteria_definitions[which(criteria_definitions$Criteria == .$Criteria[x]), .$Scope[x]]))
  
              # Number of species
  maxCol <- max(sapply(species$`Criteria met`, function(x) str_count(x, ";")))+1
  criteriaCols <- paste0("Col", 1:maxCol)
  
  criteriaInfo <- species %>%
    filter(!is.na(`Criteria met`)) %>%
    select(`Scientific name`, `Criteria met`) %>%
    separate(`Criteria met`, into=criteriaCols, sep="; ", fill="right") %>%
    pivot_longer(criteriaCols, names_to = "Remove", values_to="Criteria met") %>%
    filter(!is.na(`Criteria met`)) %>%
    group_by(`Criteria met`) %>%
    summarise(NSpecies = n()) %>%
    ungroup() %>%
    left_join(criteriaInfo, ., by=c("CriteriaFull" = "Criteria met"))
  
              # Species names
  criteriaInfo <- species %>%
    filter(!is.na(`Criteria met`)) %>%
    select(`Scientific name`, `Criteria met`) %>%
    separate(`Criteria met`, into=criteriaCols, sep="; ", fill="right") %>%
    pivot_longer(criteriaCols, names_to = "Remove", values_to="Criteria met") %>%
    filter(!is.na(`Criteria met`)) %>%
    arrange(`Scientific name`) %>%
    group_by(`Criteria met`) %>%
    summarise(speciesNames = paste(`Scientific name`, collapse=", ")) %>%
    ungroup() %>%
    left_join(criteriaInfo, ., by=c("CriteriaFull" = "Criteria met"))
  
  # PICK UP HERE
              # Flextable
  criteriaInfo_ft <- criteriaInfo %>%
    mutate(Label = "") %>%
    mutate(Blank = "") %>%
    flextable(col_keys = c("Blank", "Label")) %>%
    compose(j='Label', value=as_paragraph(as_chunk(x=paste0(as.character("\u25CF"), " ", Scope, " ", Criteria, " [criterion met by ", NSpecies, " species]", " - ", Definition, " (")), as_chunk(x=speciesNames, props=fp_text(font.size=11, font.family='Calibri', italic=T)), as_chunk(x=")."))) %>%
    font(fontname="Calibri", part="body") %>%
    fontsize(size=11, part='body') %>%
    width(j=colnames(.), width=c(0.3, 9)) %>%
    delete_part(part='header') %>%
    border_remove() %>%
    align(j=2, align = "left", part = "body")
  
        # Species assessments
              # Get information
  speciesAssessments <- species %>%
    filter(!is.na(`Criteria met`)) %>%
    mutate_at(vars(`Reproductive Units (RU)`, `Min site estimate`, `Best site estimate`, `Max site estimate`, `Min reference estimate`, `Best reference estimate`, `Max reference estimate`), as.double) %>%
    mutate(PercentAtSite = round(100 * `Best site estimate`/`Best reference estimate`, 1)) %>%
    mutate(Blank = "") %>%
    mutate(Status = ifelse(grepl("A1", `Criteria met`, fixed=T),
                           paste0(Status, " (", `Status assessment agency`, ")"),
                           "Not applicable")) %>%
    mutate(SiteEstimate_Min = as.character(`Min site estimate`),
           SiteEstimate_Best = as.character(`Best site estimate`),
           SiteEstimate_Max = as.character(`Max site estimate`),
           TotalEstimate_Min = as.character(`Min reference estimate`),
           TotalEstimate_Best = as.character(`Best reference estimate`),
           TotalEstimate_Max = as.character(`Max reference estimate`)) %>%
    mutate(AssessmentParameter = sapply(`Assessment parameter`, function(x) str_to_sentence(substr(x, start=str_locate(x, "\\)")[1,1]+2, stop=nchar(x))))) %>%
    mutate(AssessmentParameter = ifelse(AssessmentParameter %in% c("Area of occupancy", "Extent of suitable habitat", "Range"), paste(AssessmentParameter, "(km2)"), AssessmentParameter)) %>%
    select(`Scientific name`, Status, `Criteria met`, `Reproductive Units (RU)`, `RU Source`, AssessmentParameter, Blank, SiteEstimate_Min, SiteEstimate_Best, SiteEstimate_Max, `Year of site estimate`, `Derivation of best estimate`, `Explanation of site estimates`, `Sources of site estimates`, TotalEstimate_Min, TotalEstimate_Best, TotalEstimate_Max, `Explanation of reference estimates`, `Sources of reference estimates`, PercentAtSite)
  
              # Separate global and national assessments
  speciesAssessments_g <- speciesAssessments %>%
    filter(grepl("g", `Criteria met`, fixed=T)) %>%
    mutate(`Criteria met` = substr(`Criteria met`, start=2, stop=nchar(`Criteria met`)))
  
  speciesAssessments_n <- speciesAssessments %>%
    filter(grepl("n", `Criteria met`, fixed=T)) %>%
    mutate(`Criteria met` = substr(`Criteria met`, start=2, stop=nchar(`Criteria met`)))
  
  if(!(nrow(speciesAssessments_g) + nrow(speciesAssessments_n)) == nrow(speciesAssessments)){stop("Some assessments are not being correctly classified as global or national assessments. This is an error with the code. Please contact Chloé and provide her with this error message.")}
  rm(speciesAssessments)
  
              # Information for the footnotes
  footnotes_g <- speciesAssessments_g %>%
    select(`RU Source`, `Derivation of best estimate`, `Explanation of site estimates`, `Sources of site estimates`, `Explanation of reference estimates`, `Sources of reference estimates`) %>%
    mutate(`Derivation of best estimate` = sapply(`Derivation of best estimate`, function(x) ifelse(substr(x, start=nchar(x), stop=nchar(x)) == ".", x, paste0(x, ".")))) %>% 
    mutate(`Explanation of site estimates` = sapply(`Explanation of site estimates`, function(x) ifelse(substr(x, start=nchar(x), stop=nchar(x)) == ".", x, paste0(x, ".")))) %>%
    mutate(`Sources of site estimates` = sapply(`Sources of site estimates`, function(x) ifelse(substr(x, start=nchar(x), stop=nchar(x)) == ".", x, paste0(x, ".")))) %>%
    mutate(`Explanation of reference estimates` = sapply(`Explanation of reference estimates`, function(x) ifelse(substr(x, start=nchar(x), stop=nchar(x)) == ".", x, paste0(x, ".")))) %>%
    mutate(`Sources of reference estimates` = sapply(`Sources of reference estimates`, function(x) ifelse(substr(x, start=nchar(x), stop=nchar(x)) == ".", x, paste0(x, ".")))) %>%
    mutate(Site_Source = paste0("Derivation of site estimate: ", `Derivation of best estimate`, " Explanation of site estimate(s): ", `Explanation of site estimates`, " Source(s) of site estimate(s): ", `Sources of site estimates`)) %>%
    mutate(Reference_Source = paste0("Explanation of global estimate(s): ", `Explanation of reference estimates`, " Source(s) of global estimate(s): ", `Sources of reference estimates`)) %>%
    select(`RU Source`, Site_Source, Reference_Source)
  
  footnotes_n <- speciesAssessments_n %>%
    select(`RU Source`, `Derivation of best estimate`, `Explanation of site estimates`, `Sources of site estimates`, `Explanation of reference estimates`, `Sources of reference estimates`) %>%
    mutate(`Derivation of best estimate` = sapply(`Derivation of best estimate`, function(x) ifelse(substr(x, start=nchar(x), stop=nchar(x)) == ".", x, paste0(x, ".")))) %>% 
    mutate(`Explanation of site estimates` = sapply(`Explanation of site estimates`, function(x) ifelse(substr(x, start=nchar(x), stop=nchar(x)) == ".", x, paste0(x, ".")))) %>%
    mutate(`Sources of site estimates` = sapply(`Sources of site estimates`, function(x) ifelse(substr(x, start=nchar(x), stop=nchar(x)) == ".", x, paste0(x, ".")))) %>%
    mutate(`Explanation of reference estimates` = sapply(`Explanation of reference estimates`, function(x) ifelse(substr(x, start=nchar(x), stop=nchar(x)) == ".", x, paste0(x, ".")))) %>%
    mutate(`Sources of reference estimates` = sapply(`Sources of reference estimates`, function(x) ifelse(substr(x, start=nchar(x), stop=nchar(x)) == ".", x, paste0(x, ".")))) %>%
    mutate(Site_Source = paste0("Derivation of site estimate: ", `Derivation of best estimate`, " Explanation of site estimate(s): ", `Explanation of site estimates`, " Source(s) of site estimate(s): ", `Sources of site estimates`)) %>%
    mutate(Reference_Source = paste0("Explanation of national estimate(s): ", `Explanation of reference estimates`, " Source(s) of national estimate(s): ", `Sources of reference estimates`)) %>%
    select(`RU Source`, Site_Source, Reference_Source)
  
              # Information for the main table
  speciesAssessments_g %<>% select(-c(`RU Source`, `Derivation of best estimate`, `Explanation of site estimates`, `Sources of site estimates`, `Explanation of reference estimates`, `Sources of reference estimates`))
  
  speciesAssessments_n %<>% select(-c(`RU Source`, `Derivation of best estimate`, `Explanation of site estimates`, `Sources of site estimates`, `Explanation of reference estimates`, `Sources of reference estimates`))
  
              # Assess whether min/max should be retained
                    # Remove min/max values that are identical to each other
  speciesAssessments_g %<>%
    mutate(SiteEstimate_Min = ifelse(SiteEstimate_Min == SiteEstimate_Max, NA, SiteEstimate_Min),
           SiteEstimate_Max = ifelse(SiteEstimate_Min == SiteEstimate_Max, NA, SiteEstimate_Max)) %>%
    mutate(TotalEstimate_Min = ifelse(TotalEstimate_Min == TotalEstimate_Max, NA, TotalEstimate_Min),
           TotalEstimate_Max = ifelse(TotalEstimate_Min == TotalEstimate_Max, NA, TotalEstimate_Max))
    
  speciesAssessments_n %<>%
    mutate(SiteEstimate_Min = ifelse(SiteEstimate_Min == SiteEstimate_Max, NA, SiteEstimate_Min),
           SiteEstimate_Max = ifelse(SiteEstimate_Min == SiteEstimate_Max, NA, SiteEstimate_Max)) %>%
    mutate(TotalEstimate_Min = ifelse(TotalEstimate_Min == TotalEstimate_Max, NA, TotalEstimate_Min),
           TotalEstimate_Max = ifelse(TotalEstimate_Min == TotalEstimate_Max, NA, TotalEstimate_Max))
  
                    # Check if there is only a best estimate
                          # Global
  if(sum(!is.na(speciesAssessments_g$TotalEstimate_Min)) + sum(!is.na(speciesAssessments_g$TotalEstimate_Max)) + sum(!is.na(speciesAssessments_g$SiteEstimate_Min)) + sum(!is.na(speciesAssessments_g$SiteEstimate_Max)) == 0){
    bestOnly_g <- T
    speciesAssessments_g %<>% select(-c(SiteEstimate_Min, SiteEstimate_Max, TotalEstimate_Min, TotalEstimate_Max))
  }else{
    bestOnly_g <- F
  }
  
                          # National
  if(sum(!is.na(speciesAssessments_n$TotalEstimate_Min)) + sum(!is.na(speciesAssessments_n$TotalEstimate_Max)) + sum(!is.na(speciesAssessments_n$SiteEstimate_Min)) + sum(!is.na(speciesAssessments_n$SiteEstimate_Max)) == 0){
    bestOnly_n <- T
    speciesAssessments_n %<>% select(-c(SiteEstimate_Min, SiteEstimate_Max, TotalEstimate_Min, TotalEstimate_Max))
  }else{
    bestOnly_n <- F
  }
  
  # Format flextables
        # Species assessment - Global
  if(nrow(speciesAssessments_g) > 0){
    if(bestOnly_g){
      speciesAssessments_g_ft <- speciesAssessments_g %>%
        flextable() %>%
        width(j=colnames(.), width=c(1.5,1.3,0.65,1.3,1.3,0.05,0.7,0.7,0.7,0.8)) %>%
        set_header_labels(values=list(`Scientific name` = "Species", Status = "Status*", `Criteria met`="Criteria Met", `Reproductive Units (RU)` = "# of Reproductive Units", AssessmentParameter = 'Assessment Parameter', Blank='', SiteEstimate_Best = "Value", `Year of site estimate` = "Year", TotalEstimate_Best = 'Global Estimate', PercentAtSite = "% of Global Pop. at Site")) %>%
        add_header_row(values = c("Species", "Status*", "Criteria Met", "# of Reproductive Units", "Assessment Parameter", "", "Site Estimate", 'Global Estimate', "% of Global Pop. at Site"), colwidths=c(1, 1, 1, 1, 1, 1, 2, 1, 1)) %>%
        align(align = "center", part="header") %>%
        font(fontname="Calibri", part="header") %>%
        fontsize(size=11, part='header') %>%
        bold(i=1, bold=T, part='header') %>%
        merge_v(part = "header") %>%
        font(fontname="Calibri", part="body") %>%
        fontsize(size=11, part='body') %>%
        italic(j=1, part='body') %>%
        hline_top(part="all") %>%
        border_remove() %>%
        hline(border = fp_border(width = 1), part="header") %>%
        hline_top(border = fp_border(width = 2), part="header") %>%
        hline_bottom(border = fp_border(width = 2), part="header") %>%
        hline_bottom(border=fp_border(width=1), part='body') %>%
        align(j=c(2,3,4,7,8,9,10), align = "center", part = "body")
      
    }else{
      speciesAssessments_g_ft <- speciesAssessments_g %>%
        mutate(Blank2 = "") %>%
        relocate(Blank2, .after = `Year of site estimate`) %>%
        flextable() %>%
        width(j=colnames(.), width=c(1.4,1.2,0.65,1.1,0.9,0.05,0.4,0.4,0.4,0.5,0.05,0.4,0.4,0.4,0.8)) %>%
        set_header_labels(values=list(`Scientific name` = "Species", Status = "Status*", `Criteria met`="Criteria Met", `Reproductive Units (RU)` = "# of Reproductive Units", AssessmentParameter = 'Assessment Parameter', Blank='', SiteEstimate_Min = "Min", SiteEstimate_Best = "Best", SiteEstimate_Max = "Max", SiteEstimate_Year = "Year", Blank2 = "", TotalEstimate_Min = "Min", TotalEstimate_Best = "Best", TotalEstimate_Max = "Max", PercentAtSite = "% of Global Pop. at Site")) %>%
        add_header_row(values = c("Species", "Status*", "Criteria Met", "# of Reproductive Units", "Assessment Parameter", "", "Site Estimate", "", "Global Estimate", "% of Global Pop. at Site"), colwidths=c(1, 1, 1, 1, 1, 1, 4, 1, 3, 1)) %>%
        align(align = "center", part="header") %>%
        font(fontname="Calibri", part="header") %>%
        fontsize(size=11, part='header') %>%
        bold(i=1, bold=T, part='header') %>%
        merge_v(part = "header") %>%
        font(fontname="Calibri", part="body") %>%
        fontsize(size=11, part='body') %>%
        italic(j=1, part='body') %>%
        hline_top(part="all") %>%
        border_remove() %>%
        hline(border = fp_border(width = 1), part="header") %>%
        hline_top(border = fp_border(width = 2), part="header") %>%
        hline_bottom(border = fp_border(width = 2), part="header") %>%
        hline_bottom(border=fp_border(width=1), part='body') %>%
        align(j=c(2,3,4,7,8,9,10,12,13,14,15), align = "center", part = "body")
    }
  }else{
    speciesAssessments_g_ft <- ""
  }
  
        # Species assessment - National
  if(nrow(speciesAssessments_n) > 0){
    if(bestOnly_n){
      speciesAssessments_n_ft <- speciesAssessments_n %>%
        flextable() %>%
        width(j=colnames(.), width=c(1.5,1.3,0.65,1.3,1.3,0.05,0.7,0.7,0.7,0.8)) %>%
        set_header_labels(values=list(`Scientific name` = "Species", Status = "Status*", `Criteria met`="Criteria Met", `Reproductive Units (RU)` = "# of Reproductive Units", AssessmentParameter = 'Assessment Parameter', Blank='', SiteEstimate_Best = "Value", `Year of site estimate` = "Year", TotalEstimate_Best = 'National Estimate', PercentAtSite = "% of National Pop. at Site")) %>%
        add_header_row(values = c("Species", "Status*", "Criteria Met", "# of Reproductive Units", "Assessment Parameter", "", "Site Estimate", 'National Estimate', "% of National Pop. at Site"), colwidths=c(1, 1, 1, 1, 1, 1, 2, 1, 1)) %>%
        align(align = "center", part="header") %>%
        font(fontname="Calibri", part="header") %>%
        fontsize(size=11, part='header') %>%
        bold(i=1, bold=T, part='header') %>%
        merge_v(part = "header") %>%
        font(fontname="Calibri", part="body") %>%
        fontsize(size=11, part='body') %>%
        italic(j=1, part='body') %>%
        hline_top(part="all") %>%
        border_remove() %>%
        hline(border = fp_border(width = 1), part="header") %>%
        hline_top(border = fp_border(width = 2), part="header") %>%
        hline_bottom(border = fp_border(width = 2), part="header") %>%
        hline_bottom(border=fp_border(width=1), part='body') %>%
        align(j=c(2,3,4,7,8,9,10), align = "center", part = "body")
      
    }else{
      speciesAssessments_n_ft <- speciesAssessments_n %>%
        mutate(Blank2 = "") %>%
        relocate(Blank2, .after = `Year of site estimate`) %>%
        flextable() %>%
        width(j=colnames(.), width=c(1.4,1.2,0.65,1.1,0.9,0.05,0.4,0.4,0.4,0.5,0.05,0.4,0.4,0.4,0.8)) %>%
        set_header_labels(values=list(`Scientific name` = "Species", Status = "Status*", `Criteria met`="Criteria Met", `Reproductive Units (RU)` = "# of Reproductive Units", AssessmentParameter = 'Assessment Parameter', Blank='', SiteEstimate_Min = "Min", SiteEstimate_Best = "Best", SiteEstimate_Max = "Max", SiteEstimate_Year = "Year", Blank2 = "", TotalEstimate_Min = "Min", TotalEstimate_Best = "Best", TotalEstimate_Max = "Max", PercentAtSite = "% of National Pop. at Site")) %>%
        add_header_row(values = c("Species", "Status*", "Criteria Met", "# of Reproductive Units", "Assessment Parameter", "", "Site Estimate", "", "National Estimate", "% of National Pop. at Site"), colwidths=c(1, 1, 1, 1, 1, 1, 4, 1, 3, 1)) %>%
        align(align = "center", part="header") %>%
        font(fontname="Calibri", part="header") %>%
        fontsize(size=11, part='header') %>%
        bold(i=1, bold=T, part='header') %>%
        merge_v(part = "header") %>%
        font(fontname="Calibri", part="body") %>%
        fontsize(size=11, part='body') %>%
        italic(j=1, part='body') %>%
        hline_top(part="all") %>%
        border_remove() %>%
        hline(border = fp_border(width = 1), part="header") %>%
        hline_top(border = fp_border(width = 2), part="header") %>%
        hline_bottom(border = fp_border(width = 2), part="header") %>%
        hline_bottom(border=fp_border(width=1), part='body') %>%
        align(j=c(2,3,4,7,8,9,10,12,13,14,15), align = "center", part = "body")
    }
  }else{
    speciesAssessments_n_ft <- ""
  }
  
  # Add footnotes, with formatted hyperlinks
        # Global
  if(nrow(speciesAssessments_g) > 0){
    footnote <- 0
    for(i in 1:nrow(speciesAssessments_g)){
      col <- which(grepl("http", footnotes_n[i,]), arr.ind = TRUE)
      
      for(c in 1:ncol(footnotes_g)){
        string <- footnotes_g[i,c]
        
        if(!is.na(string)){
          footnote <- footnote+1
          
          # If there's a link in the footnote
          if(c %in% col){
            urls <- str_locate_all(string, "http")[[1]][,1]
            urlIDs <- paste0("url", urls)
            spaces <- str_locate_all(string, " ")[[1]][,1] %>%
              ifelse(length(.) == 0, -1, .)
            links <- list()
            
            for(u in 1:length(urls)){
              url <- urls[u]
              
              if(spaces[length(spaces)] > url){
                space <- spaces[which(spaces > url)][1]
                link <- substr(string, start=url, stop=space-1)
                
              }else{
                link <- substr(string, start=url, stop=nchar(string))
              }
              
              # Remove full-stops and parentheses at the end
                    # First round
              if(substr(link, start=nchar(link), stop=nchar(link)) == "."){
                link <- substr(link, start=1, stop=nchar(link)-1)
                
              }else if(substr(link, start=nchar(link), stop=nchar(link)) == ")"){
                link <- substr(link, start=1, stop=nchar(link)-1)
              }
              
                    # Second round
              if(substr(link, start=nchar(link), stop=nchar(link)) == "."){
                link <- substr(link, start=1, stop=nchar(link)-1)
                
              }else if(substr(link, start=nchar(link), stop=nchar(link)) == ")"){
                link <- substr(link, start=1, stop=nchar(link)-1)
              }
              
              links[[urlIDs[u]]] <- link
            }
            
            # Create call
            call_substr <- rep("substr", each=length(urls)+1)
            call_hyperlink <- rep("hyperlink", each=length(urls))
            call_all <- c(sapply(seq_along(call_substr), function(i) append(call_substr[i], call_hyperlink[i], i)))
            call_all <- call_all[which(!is.na(call_all))]
            start <- 1
            
            for(call in 1:length(call_all)){
              if(call_all[call] == "substr"){
                
                if(call == 1){
                  text <- paste0("substr(string, start=", start, ", stop=urls[", call, "]-1)")
                }else if(!call == length(call_all)){
                  text <- paste0("substr(string, start=", start, ", stop=urls[", (call+1)/2, "]-1)")
                }else{
                  text <- paste0("substr(string, start=", start, ", stop=nchar(string))")
                }
                
              }else{
                text <- paste0("hyperlink_text(x='link', url=links[", call/2, "], props = fp_text(color='blue', font.size=11, underlined=T, font.family = 'Calibri'))")
                start <- urls[call/2] + nchar(links[call/2])
              }
              
              if(call == 1){
                call_final <- text
              }else{
                call_final <- paste(call_final, text, sep=', ')
              }
            }
            
            if(bestOnly){
              call_final <- paste0("elements_ft %<>% footnote(i=i, j=ifelse(c==1, 4, ifelse(c==2, 7, 9)), value=as_paragraph(", call_final,"), ref_symbols=as.integer(footnote))")
            }else{
              call_final <- paste0("elements_ft %<>% footnote(i=i, j=ifelse(c==1, 4, ifelse(c==2, 8, 13)), value=as_paragraph(", call_final,"), ref_symbols=as.integer(footnote))")
            }
            
            # Evaluate call
            eval(parse(text=call_final))
            
            # If there is no link in the footnote
          }else{
            
            if(bestOnly_g){
              speciesAssessments_g_ft %<>% footnote(i=i, j=ifelse(c==1, 4, ifelse(c==2, 7, 9)), value=as_paragraph(as.character(string)), ref_symbols=as.integer(footnote))
            }else{
              speciesAssessments_g_ft %<>% footnote(i=i, j=ifelse(c==1, 4, ifelse(c==2, 8, 13)), value=as_paragraph(as.character(string)), ref_symbols=as.integer(footnote))
            }
          }
        }
      }
    }
  }
 
        # National
  if(nrow(speciesAssessments_n) > 0){
    footnote <- 0
    for(i in 1:nrow(speciesAssessments_n)){
      col <- which(grepl("http", footnotes_n[i,]), arr.ind = TRUE)
      
      for(c in 1:ncol(footnotes_n)){
        string <- footnotes_n[i,c]
        
        if(!is.na(string)){
          footnote <- footnote+1
          
          # If there's a link in the footnote
          if(c %in% col){
            urls <- str_locate_all(string, "http")[[1]][,1]
            urlIDs <- paste0("url", urls)
            spaces <- str_locate_all(string, " ")[[1]][,1] %>%
              ifelse(length(.) == 0, -1, .)
            links <- list()
            
            for(u in 1:length(urls)){
              url <- urls[u]
              
              if(spaces[length(spaces)] > url){
                space <- spaces[which(spaces > url)][1]
                link <- substr(string, start=url, stop=space-1)
                
              }else{
                link <- substr(string, start=url, stop=nchar(string))
              }
              
              # Remove full-stops and parentheses at the end
              # First round
              if(substr(link, start=nchar(link), stop=nchar(link)) == "."){
                link <- substr(link, start=1, stop=nchar(link)-1)
                
              }else if(substr(link, start=nchar(link), stop=nchar(link)) == ")"){
                link <- substr(link, start=1, stop=nchar(link)-1)
              }
              
              # Second round
              if(substr(link, start=nchar(link), stop=nchar(link)) == "."){
                link <- substr(link, start=1, stop=nchar(link)-1)
                
              }else if(substr(link, start=nchar(link), stop=nchar(link)) == ")"){
                link <- substr(link, start=1, stop=nchar(link)-1)
              }
              
              links[[urlIDs[u]]] <- link
            }
            
            # Create call
            call_substr <- rep("substr", each=length(urls)+1)
            call_hyperlink <- rep("hyperlink", each=length(urls))
            call_all <- c(sapply(seq_along(call_substr), function(i) append(call_substr[i], call_hyperlink[i], i)))
            call_all <- call_all[which(!is.na(call_all))]
            start <- 1
            
            for(call in 1:length(call_all)){
              if(call_all[call] == "substr"){
                
                if(call == 1){
                  text <- paste0("substr(string, start=", start, ", stop=urls[", call, "]-1)")
                }else if(!call == length(call_all)){
                  text <- paste0("substr(string, start=", start, ", stop=urls[", (call+1)/2, "]-1)")
                }else{
                  text <- paste0("substr(string, start=", start, ", stop=nchar(string))")
                }
                
              }else{
                text <- paste0("hyperlink_text(x='link', url=links[", call/2, "], props = fp_text(color='blue', font.size=11, underlined=T, font.family = 'Calibri'))")
                start <- urls[call/2] + nchar(links[call/2])
              }
              
              if(call == 1){
                call_final <- text
              }else{
                call_final <- paste(call_final, text, sep=', ')
              }
            }
            
            if(bestOnly){
              call_final <- paste0("elements_ft %<>% footnote(i=i, j=ifelse(c==1, 4, ifelse(c==2, 7, 9)), value=as_paragraph(", call_final,"), ref_symbols=as.integer(footnote))")
            }else{
              call_final <- paste0("elements_ft %<>% footnote(i=i, j=ifelse(c==1, 4, ifelse(c==2, 8, 13)), value=as_paragraph(", call_final,"), ref_symbols=as.integer(footnote))")
            }
            
            # Evaluate call
            eval(parse(text=call_final))
            
            # If there is no link in the footnote
          }else{
            
            if(bestOnly_n){
              speciesAssessments_n_ft %<>% footnote(i=i, j=ifelse(c==1, 4, ifelse(c==2, 7, 9)), value=as_paragraph(as.character(string)), ref_symbols=as.integer(footnote))
            }else{
              speciesAssessments_n_ft %<>% footnote(i=i, j=ifelse(c==1, 4, ifelse(c==2, 8, 13)), value=as_paragraph(as.character(string)), ref_symbols=as.integer(footnote))
            }
          }
        }
      }
    }
  }
  
  # Add padding
        # Global
  if(nrow(speciesAssessments_g) > 0){
    speciesAssessments_g_ft %<>%
      padding(padding.top = 10, part='footer') %>%
      font(fontname='Calibri', part='footer')
  }
  
        # National
  if(nrow(speciesAssessments_n) > 0){
    speciesAssessments_n_ft %<>%
      padding(padding.top = 10, part='footer') %>%
      font(fontname='Calibri', part='footer')
  }
  
  # Prepare final tables
        # Global
  if(nrow(speciesAssessments_g) > 0){
    elementsOnly_g <- speciesAssessments_g_ft %>%
      delete_part(part='footer')
    
    footnotesOnly_g <- speciesAssessments_g_ft %>%
      delete_part(part='header') %>%
      delete_part(part='body') %>%
      bg(bg = "#EFEFEF", part = "footer")
  }
  
        # National
  if(nrow(speciesAssessments_n) > 0){
    elementsOnly_n <- speciesAssessments_n_ft %>%
      delete_part(part='footer')
    
    footnotesOnly_n <- speciesAssessments_n_ft %>%
      delete_part(part='header') %>%
      delete_part(part='body') %>%
      bg(bg = "#EFEFEF", part = "footer")
  }
  
  # Trigger Elements summary
  elementsSummary <- species %>%
    filter(!is.na(`Criteria met`)) %>%
    select(`Common name`, `Scientific name`) %>%
    unique() %>%
    pivot_longer(., cols=c(`Common name`, `Scientific name`), names_to="Type") %>%
    select(-Type) %>%
    t() %>%
    data.frame() %>%
    mutate(Prefix = paste0(as.character("\u25CF"), " Species: "))
  elementsSummary <- elementsSummary[,c(ncol(elementsSummary), 1:(ncol(elementsSummary)-1))]
  
  elementsSummary_ft <- flextable(elementsSummary, col_keys = c("Blank", "Label"), defaults=list(fontname="Calibri", font.size=11)) %>%
    width(j=colnames(.), width=c(0.3, 9))
  
  extraCall <- ""
  if(ncol(elementsSummary) > 3){
    
    # Keep only columns with common names
    spp <- 4:ncol(elementsSummary)
    spp <- spp[lapply(spp, "%%", 2) == 0]
    
    for(i in spp){
      extraCall <- paste0(extraCall, ", as_chunk(x=', '), as_chunk(x=X", i-1, "), as_chunk(x=' ('), as_chunk(x=X", i, ", props=fp_text(font.size=11, font.family='Calibri', italic = T)), as_chunk(x=')')")
    }
  }
  compose_call <- paste0("elementsSummary_ft %<>% compose(j='Label', value=as_paragraph(as_chunk(Prefix), as_chunk(x=X1), as_chunk(x=' ('), as_chunk(x=X2, props=fp_text(font.size=11, font.family='Calibri', italic = T)), as_chunk(x=')')", 
                         extraCall,
                         "))")
  eval(parse(text=compose_call))
  
  elementsSummary_ft %<>%
    delete_part(part='header') %>%
    border_remove() %>%
    align(j=2, align = "left", part = "body") %>%
    font(fontname = "Calibri", part="body")
  
  # Subtitle (cover page)
  elementsSummary %<>% select(-Prefix)
  subtitle_ft <- flextable(elementsSummary, col_keys = "Label", defaults=list(fontname="Calibri", font.size=12, color='#5A5A5A')) %>%
    width(j=colnames(.), width=c(9))
  
  extraCall <- ""
  if(ncol(elementsSummary) > 3){
    
    # Keep only columns with common names
    spp <- 4:ncol(elementsSummary)
    spp <- spp[lapply(spp, "%%", 2) == 0]
    
    for(i in spp){
      extraCall <- paste0(extraCall, ", as_chunk(x=', ', props=fp_text(font.size=12, font.family='Calibri', color='#5A5A5A')), as_chunk(x=X", i-1, ", props=fp_text(font.size=12, font.family='Calibri', color='#5A5A5A')), as_chunk(x=' (', props=fp_text(font.size=12, font.family='Calibri', color='#5A5A5A')), as_chunk(x=X", i, ", props=fp_text(font.size=12, font.family='Calibri', italic = T, color='#5A5A5A')), as_chunk(x=')', props=fp_text(font.size=12, font.family='Calibri', color='#5A5A5A'))")
    }
  }
  compose_call <- paste0("subtitle_ft %<>% compose(j='Label', value=as_paragraph(as_chunk(x=X1, props=fp_text(font.size=12, font.family='Calibri', color='#5A5A5A')), as_chunk(x=' (', props=fp_text(font.size=12, font.family='Calibri', color='#5A5A5A')), as_chunk(x=X2, props=fp_text(font.size=12, font.family='Calibri', italic = T, color='#5A5A5A')), as_chunk(x=')', props=fp_text(font.size=12, font.family='Calibri', color='#5A5A5A'))", 
                         extraCall,
                         "))")
  eval(parse(text=compose_call))
  
  subtitle_ft %<>%
    delete_part(part='header') %>%
    border_remove() %>%
    align(j=1, align = "left", part = "body") %>%
    fontsize(size=12, part='body')
  
  # Technical Review
  technicalReview_ft <- technicalReview %>%
    flextable() %>%
    width(j=colnames(.), width=c(1.4,2,2,3.6)) %>%
    align(align = "center", part="header") %>%
    font(fontname="Calibri", part="header") %>%
    fontsize(size=11, part='header') %>%
    bold(i=1, bold=T, part='header') %>%
    merge_v(part = "header") %>%
    font(fontname="Calibri", part="body") %>%
    fontsize(size=11, part='body') %>%
    hline_top(part="all")
  
  # General Review
  generalReview_ft <- generalReview %>%
    flextable() %>%
    width(j=colnames(.), width=c(2.4,3.6,3)) %>%
    align(align = "center", part="header") %>%
    font(fontname="Calibri", part="header") %>%
    fontsize(size=11, part='header') %>%
    bold(i=1, bold=T, part='header') %>%
    merge_v(part = "header") %>%
    font(fontname="Calibri", part="body") %>%
    fontsize(size=11, part='body') %>%
    hline_top(part="all")
  
  # Additional Site Information
  additionalInfo <- data.frame(Type = character(),
                               Value = character(),
                               stringsAsFactors = F)
  
        # Nomination rationale
  additionalInfo[1, ] <- c("Rationale for site nomination", nominationRationale)
  
        # Assessed elements that did not meet KBA criteria
  speciesNotTriggers <- species %>%
    filter(is.na(`Criteria met`)) %>%
    pull(`Scientific name`) %>%
    unique() %>%
    paste(., collapse=", ")

  additionalInfo[2, ] <- c("Biodiversity elements that were assessed but did not meet KBA criteria", ifelse(speciesNotTriggers == "", "-", speciesNotTriggers))
  
        # Additional biodiversity
  additionalInfo[3, ] <- c("Additional biodiversity at the site", ifelse(is.na(additionalBiodiversity), "-", additionalBiodiversity))
  
        # Percent protected
  additionalInfo[4, ] <- c("Percent of site covered by protected areas", percentProtected)
  
        # Customary jurisdiction at site
  additionalInfo[5, ] <- c("Customary jurisdiction at site", ifelse(is.na(customaryJurisdiction), "-", customaryJurisdiction))
  
        # Ongoing conservation actions
  ongoingActions <- actions %>%
    filter(Ongoing == "TRUE") %>%
    pull(Action) %>%
    substr(., start=5, stop=nchar(.)) %>%
    paste(., collapse="; ")
  
  additionalInfo[6, ] <- c("Ongoing conservation actions", ifelse((length(ongoingActions) == 0) | (ongoingActions == ""), "None", ongoingActions))
  
        # Ongoing threats
  if(!noThreats){
    threatText <- threats %>%
      pull(`Level 1`) %>%
      unique() %>%
      substr(., start=3, stop=nchar(.)) %>%
      trimws() %>%
      sort() %>%
      paste(., collapse='; ')
  }else{
    threatText <- "-"
  }
  
  additionalInfo[7, ] <- c("Ongoing threats", threatText)
  
        # Conservation actions needed
  neededActions <- actions %>%
    filter(Needed == "TRUE") %>%
    pull(Action) %>%
    substr(., start=5, stop=nchar(.)) %>%
    paste(., collapse="; ")
  
  additionalInfo[8, ] <- c("Conservation actions needed", ifelse((length(neededActions) == 0) | (neededActions == ""), "-", neededActions))
  
        # Make it a flextable
  additionalInfo_ft <- additionalInfo %>%
    flextable() %>%
    width(j=colnames(.), width=c(3,6)) %>%
    font(fontname="Calibri", part="body") %>%
    fontsize(size=11, part='body') %>%
    italic(i=2, j=2, part='body') %>%
    delete_part(part="header") %>%
    theme_zebra() %>%
    align(j=c(1,2), align='left', part='body') %>%
    bold(j=1)
  
  # Citations
  if(nrow(citations) == 0){
    citations[1,] <- c("", "No references provided.", "", "")
  }
  
  citations_ft <- citations %>%
    arrange(`Long Citation`) %>%
    select(`Long Citation`) %>%
    flextable() %>%
    delete_part(part = "header") %>%
    width(j=colnames(.), width=9) %>%
    border_remove() %>%
    font(fontname="Calibri", part="body")
  
  # List all flextables
  if(scope == "Global and National"){
    FT <- list(subtitle = subtitle_ft, triggerElements = elementsSummary_ft, criteria = criteriaInfo_ft, elements_g = elementsOnly_g, elementFootnotes_g = footnotesOnly_g, elements_n = elementsOnly_n, elementFootnotes_n = footnotesOnly_n, technicalReview = technicalReview_ft, generalReview = generalReview_ft, additionalInfo = additionalInfo_ft, references = citations_ft)
  }else if(scope == "Global"){
    FT <- list(subtitle = subtitle_ft, triggerElements = elementsSummary_ft, criteria = criteriaInfo_ft, elements_g = elementsOnly_g, elementFootnotes_g = footnotesOnly_g, technicalReview = technicalReview_ft, generalReview = generalReview_ft, additionalInfo = additionalInfo_ft, references = citations_ft)
  }else{
    FT <- list(subtitle = subtitle_ft, triggerElements = elementsSummary_ft, criteria = criteriaInfo_ft, elements_n = elementsOnly_n, elementFootnotes_n = footnotesOnly_n, technicalReview = technicalReview_ft, generalReview = generalReview_ft, additionalInfo = additionalInfo_ft, references = citations_ft)
  }
  
  #### Save KBA summary
  # Get template
  if(includeQuestions){
    if(includeReviewDetails){
      if(scope == "Global and National"){
        template <- 'joint_files/KBASummary_Template_NewForm_Questions_Review_GlobalNational.docx'
        ## Google Drive: https://docs.google.com/document/d/11EtnJuLgEUfudzDPhpDMNUvKPHGvRgCe/edit?usp=sharing&ouid=104844399648613391324&rtpof=true&sd=true
      }else if(scope == "Global"){
        template <- 'joint_files/KBASummary_Template_NewForm_Questions_Review_Global.docx'
        ## Google Drive: https://docs.google.com/document/d/11IxNB0isZicHfZ9L6zPwxT-AfQk0AzM1/edit?usp=sharing&ouid=104844399648613391324&rtpof=true&sd=true
      }else{
        template <- 'joint_files/KBASummary_Template_NewForm_Questions_Review_National.docx'
        ## Google Drive: https://docs.google.com/document/d/11C8_DGI7RvmgyLh7z9iJNQ8ctiZWmzg3/edit?usp=sharing&ouid=104844399648613391324&rtpof=true&sd=true
      }
    }else{
      if(scope == "Global and National"){
        template <- 'joint_files/KBASummary_Template_NewForm_Questions_NoReview_GlobalNational.docx'
        ## Google Drive: https://docs.google.com/document/d/11RfjVzkFhYGEddAMJZwxj0U5se2kqLOD/edit?usp=sharing&ouid=104844399648613391324&rtpof=true&sd=true
      }else if(scope == "Global"){
        template <- 'joint_files/KBASummary_Template_NewForm_Questions_NoReview_Global.docx'
        ## Google Drive: https://docs.google.com/document/d/116c7UuaT7MGAXGoKnfgZ8G7GyPqc_Hsv/edit?usp=sharing&ouid=104844399648613391324&rtpof=true&sd=true
      }else{
        template <- 'joint_files/KBASummary_Template_NewForm_Questions_NoReview_National.docx'
        ## Google Drive: https://docs.google.com/document/d/11NT6kSksHvmw6Kfn7PgD38rWtsHn_5RR/edit?usp=sharing&ouid=104844399648613391324&rtpof=true&sd=true
      }
    }
    
  }else{
    if(includeReviewDetails){
      if(scope == "Global and National"){
        template <- 'joint_files/KBASummary_Template_NewForm_NoQuestions_Review_GlobalNational.docx'
        ## Google Drive: https://docs.google.com/document/d/1ztHExERMAN6GfgHeu1y2jwI7PPfuspjf/edit?usp=sharing&ouid=104844399648613391324&rtpof=true&sd=true
      }else if(scope == "Global"){
        template <- 'joint_files/KBASummary_Template_NewForm_NoQuestions_Review_Global.docx'
        ## Google Drive: https://docs.google.com/document/d/1zxKFrxZjkc6VpNdBdkt80VSM5jg3-zXm/edit?usp=sharing&ouid=104844399648613391324&rtpof=true&sd=true
      }else{
        template <- 'joint_files/KBASummary_Template_NewForm_NoQuestions_Review_National.docx'
        ## Google Drive: https://docs.google.com/document/d/1zzD8vb0X8kq2_B_lXwhoxqxcj8JK9IIe/edit?usp=sharing&ouid=104844399648613391324&rtpof=true&sd=true
      }
    }else{
      if(scope == "Global and National"){
        template <- 'joint_files/KBASummary_Template_NewForm_NoQuestions_NoReview_GlobalNational.docx'
        ## Google Drive: https://docs.google.com/document/d/1--Qh4Dif9Cr8RNS9u1ODcsVvXEDLBIEG/edit?usp=sharing&ouid=104844399648613391324&rtpof=true&sd=true
      }else if(scope == "Global"){
        template <- 'joint_files/KBASummary_Template_NewForm_NoQuestions_NoReview_Global.docx'
        ## Google Drive: https://docs.google.com/document/d/1-31LLlC09UpJeH6fKFFLagPtZG8jxkzT/edit?usp=sharing&ouid=104844399648613391324&rtpof=true&sd=true
      }else{
        template <- 'joint_files/KBASummary_Template_NewForm_NoQuestions_NoReview_National.docx'
        ## Google Drive: https://docs.google.com/document/d/1mjDJVcLVkYGpc961QApZNU7YvuN4RqJc/edit?usp=sharing&ouid=104844399648613391324&rtpof=true&sd=true
      }
    }
  }
  
   #Compute document name
   doc <- paste0("Summary_", str_replace_all(string=nationalName, pattern=c(":| |\\(|\\)"), repl=""), "_", Sys.Date(), ".docx")
 
  # Save
  doc <- renderInlineCode(template, doc)
  Sys.sleep(10)
  doc <- body_add_flextables(doc, doc, FT)

  KBAforms[step] <- doc

}
return(KBAforms)
}


#form_conversion("proposal/test.xlsm", includeQuestions = FALSE, includeReviewDetails = FALSE)
#KBAforms = "proposal/test.xlsm"
#KBAforms = "proposal/KBAProposal_Ojibway Prairie Complex and Greater Park Ecosystem- 08.27.2021-RR.xlsm"
