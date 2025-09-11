summary <- function(KBAforms, reviewStage, language, app){
  
  # Options
  options(scipen = 999)
  
  # Load functions
  source_url("https://github.com/chloedebyser/KBA-Public/blob/main/KBA%20Functions.R?raw=TRUE")
  
  # Load crosswalks
        # Assessment Parameter
  if(language == "french"){
    googledrive::drive_download("https://docs.google.com/spreadsheets/d/1Tdfbakn1CHnOhzvdlqi2mq48QblTbDDP", overwrite = TRUE)
    xwalk_assessmentParameter <- read.xlsx("AssessmentParameter.xlsx")
  }
  
        # Conservation Action
  if(language == "french"){
    googledrive::drive_download("https://docs.google.com/spreadsheets/d/1TcnYlIBNOkRdhdCqW325WSwHWEWrWZtQ", overwrite = TRUE)
    xwalk_conservationAction <- read.xlsx("ConservationAction.xlsx")
  }
  
        # Criterion
  googledrive::drive_download("https://docs.google.com/spreadsheets/d/1TB8LJvQNZd2OhBmSXzc2vPsICSnyYaaS", overwrite = TRUE)
  xwalk_criterion <- read.xlsx("Criterion.xlsx")
  
        # Derivation Of Best Estimate
  if(language == "french"){
    googledrive::drive_download("https://docs.google.com/spreadsheets/d/1ToAUYfdcM1A_uI8hevOQGvf1vt8LN3TS", overwrite = TRUE)
    xwalk_derivationOfBestEstimate <- read.xlsx("DerivationOfBestEstimate.xlsx")
  }
  
        # Jurisdiction
  if(language == "french"){
    googledrive::drive_download("https://docs.google.com/spreadsheets/d/1UAouhO7S2ojZrkrbIFWZHXwyFKeDcQbE", overwrite = TRUE)
    xwalk_jurisdiction <- read.xlsx("Jurisdiction.xlsx")
  }
  
        # Threat
  if(language == "french"){
    googledrive::drive_download("https://docs.google.com/spreadsheets/d/1TpTyNQC4J_wdpgRtvdSBu1m899t9MsN9", overwrite = TRUE)
    xwalk_threat <- read.xlsx("Threat.xlsx")
  }
  
  # Prepare DB_BIOTICS_ELEMENT_NATIONAL
        # Load species list
  googledrive::drive_download("https://docs.google.com/spreadsheets/d/1R2ILLvyGMqRL8S9pfZdYIeBKXlyzckKQ", overwrite = T)
  masterSpeciesList <- read_excel("Ref_Species.xlsx", sheet=2)
  write_excel_csv(masterSpeciesList, file="Ref_Species.csv")
  masterSpeciesList <- read_csv("Ref_Species.csv")
  
        # Convert species list to BIOTICS_ELEMENT_NATIONAL
  DB_BIOTICS_ELEMENT_NATIONAL <- masterSpeciesList %>%
    {.[, c("SpeciesID", colnames(.)[which(!str_detect(colnames(.), "[[:lower:]]"))])]}
  colnames(DB_BIOTICS_ELEMENT_NATIONAL) <- tolower(colnames(DB_BIOTICS_ELEMENT_NATIONAL))
  DB_BIOTICS_ELEMENT_NATIONAL <<- DB_BIOTICS_ELEMENT_NATIONAL
  
  # Prepare DB_BIOTICS_ECOSYSTEM
        # Load ecosystem list
  googledrive::drive_download("https://docs.google.com/spreadsheets/d/1BhmzLVFZIs-SzFaH3C3q1RM8ZHZBuJCl", overwrite = T)
  masterEcosystemList <- read_excel("Ref_Ecosystems.xlsx", sheet=2)
  write_excel_csv(masterEcosystemList, file="Ref_Ecosystems.csv")
  masterEcosystemList <- read_csv("Ref_Ecosystems.csv")
  
        # Convert ecosystem list to BIOTICS_ECOSYSTEM
  DB_BIOTICS_ECOSYSTEM <- masterEcosystemList %>%
    {.[, c("EcosystemID", colnames(.)[which(!str_detect(colnames(.), "[[:lower:]]"))])]}
  colnames(DB_BIOTICS_ECOSYSTEM) <- tolower(colnames(DB_BIOTICS_ECOSYSTEM))
  DB_BIOTICS_ECOSYSTEM <<- DB_BIOTICS_ECOSYSTEM
  
  # Create a dataframe to store the success/failure state of each conversion
  convert_res <- data.frame(matrix(ncol=3))
  colnames(convert_res) <- c("Name","Result","Message")
  
  # Get list of variables to retain
  toRetain <- ls()
  
  #### Prepare the Summary(ies) ####
  for(step in 1:length(KBAforms)){
    
    rm(list=setdiff(ls(), c(toRetain, "step", "toRetain")))
    
    if(!grepl(".xlsm", KBAforms[step], fixed =  TRUE)){
      convert_res[step,"Result"] <- emo::ji("prohibited")
      convert_res[step, "Message"] <- paste(KBAforms[step], "is not a KBA proposal form")
      KBAforms[step] <- NA
      next
    }
    
    if(app){
      incProgress(1/length(KBAforms), detail = paste("form number ", step))
    }
    
    success <- FALSE # Set success to FALSE
    
    # Check that the form is a KBA Canada Proposal Form
          # Load full workbook
    wb <- loadWorkbook(KBAforms[step])
    
          # Check that the correct tabs are present
    if(sum(c("HOME", "1. PROPOSER", "2. SITE", "3. SPECIES","4. ECOSYSTEMS & C", "5. THREATS", "6. REVIEW", "7. CITATIONS", "8. CHECK") %in% getSheetNames(KBAforms[step])) != 9) {
      convert_res[step, "Result"] <- emo::ji("prohibited")
      convert_res[step, "Message"] <- paste(KBAforms[step], "is not a KBA Canada proposal form. If you need a summary for a Legacy single-site form, contact Chloé.")
      KBAforms[step] <- NA
      next
    }
    rm(wb)
    
    # Load KBA Canada Proposal Form
    read_KBACanadaProposalForm(formPath = KBAforms[step], final = ifelse(reviewStage == "steering", T, F))
    
    # Handle the site name
          # Get the name
    name <- PF_site %>%
      filter(Field == "National name") %>%
      pull(GENERAL)
    
          # Check that the name exists
    if(is.na(name)){
      convert_res[step,"Result"] <- emo::ji("prohibited")
      convert_res[step,"Message"] <- "KBA site must have a name."
      KBAforms[step] <- NA
      next
    }
    
          # Assign the name of the site to the name in result table
    convert_res[step,"Name"] <- name
    
          # Check that the name is not too long
    if(nchar(name)>80){
      convert_res[step,"Result"] <- emo::ji("prohibited")
      convert_res[step,"Message"] <- "KBA name is too long."
      KBAforms[step] <- NA
      next
    }
    
          # Check that the name does not contain any paragraph symbols
    if(grepl("\n", name, fixed=T)){
      convert_res[step,"Result"] <- emo::ji("prohibited")
      convert_res[step,"Message"] <- "KBA name should not include paragraph breaks."
      KBAforms[step] <- NA
      next
    }
    
    # Check form version compatibility
    if(!PF_formVersion %in% c(1, 1.1, 1.2)){
      convert_res[step,"Result"] <- emo::ji("prohibited")
      convert_res[step,"Message"] <- "Form version not supported. Please contact Chloé and provide her with this error message."
      KBAforms[step] <- NA
      next
    }
    
    # Format the sheets
          # 3. SPECIES
    if(nrow(PF_species) > 0){
      
                # If French is requested, translate the derivation of best estimate to French
      if(language == "french"){
        PF_species %<>%
          left_join(., xwalk_derivationOfBestEstimate, by=c("Derivation of best estimate" = "DerivationOfBestEstimate_EN")) %>%
          mutate(`Derivation of best estimate` = DerivationOfBestEstimate_FR) %>%
          select(-DerivationOfBestEstimate_FR)
      }
      
                # If French is requested, translate the species common names to French
      if(language == "french"){
        
        PF_species %<>%
          left_join(., masterSpeciesList[,c("ELEMENT_CODE", "NATIONAL_FR_NAME")], by=c("NatureServe Element Code" = "ELEMENT_CODE")) %>%
          mutate(`Common name` = NATIONAL_FR_NAME) %>%
          select(-NATIONAL_FR_NAME)
        
        if(sum(is.na(PF_species$`Common name`)) > 0){
          
          if(!sum(PF_species$`NatureServe Element Code` %in% masterSpeciesList$ELEMENT_CODE) == nrow(PF_species)){
            convert_res[step,"Result"] <- emo::ji("prohibited")
            convert_res[step,"Message"] <- "Some values for NatureServe Element Code (3. SPECIES tab) are not recognized. Please cross-check your entries with the master species list."
            KBAforms[step] <- NA
            next
            
          }else{
            convert_res[step,"Result"] <- emo::ji("prohibited")
            convert_res[step,"Message"] <- "Some species do not have a French common name in the master species list. Please contact Chloé and provide this error message."
            KBAforms[step] <- NA
            next
          }
        }
      }
      
                # If two common names are provided, only keep the first
      for(i in 1:nrow(PF_species)){
        
        if(grepl(";", PF_species$`Common name`[i])){
          PF_species$`Common name`[i] %<>% substr(., start=1, stop=unlist(gregexpr(";", PF_species$`Common name`[i]))-1)
        }
      }
      
                # Only retain information in the desired language
      error <- F
      
      for(i in 1:nrow(PF_species)){
        
        for(j in 1:ncol(PF_species)){
          
          PF_species[i,j] <- gsub("FRANÇAIS -", "FRANCAIS -", PF_species[i,j])
          PF_species[i,j] <- gsub("Français -", "FRANCAIS -", PF_species[i,j])
          PF_species[i,j] <- gsub("français -", "FRANCAIS -", PF_species[i,j])
          
          if(grepl("FRANCAIS", PF_species[i,j]) | grepl("ENGLISH", PF_species[i,j])){
            
            # Initiate language check
            checkFR <- F
            checkEN <- F
            
            # Get index of FRANCAIS annotation
            if(grepl("FRANCAIS", PF_species[i,j])){
              checkFR <- T
              startFR <- unlist(gregexpr("FRANCAIS", PF_species[i,j]))
            }
            
            # Get index of ENGLISH annotation
            if(grepl("ENGLISH", PF_species[i,j])){
              checkEN <- T
              startEN <- unlist(gregexpr("ENGLISH", PF_species[i,j]))
            }
            
            # Get desired text
            if(checkFR & checkEN){
              
              if(startFR < startEN){
                FR <- substr(PF_species[i,j], start=startFR + nchar("FRANCAIS"), stop=startEN-1)
                EN <- substr(PF_species[i,j], start=startEN + nchar("ENGLISH"), stop=nchar(PF_species[i,j]))
                
              }else{
                FR <- substr(PF_species[i,j], start=startFR + nchar("FRANCAIS"), stop=nchar(PF_species[i,j]))
                EN <- substr(PF_species[i,j], start=startEN + nchar("ENGLISH"), stop=startFR-1)
              }
              
              if(language == "english"){
                final <- EN
                
              }else{
                final <- FR
              }
              
            }else if(checkFR){
              
              if((language == "french") | (reviewStage == "steering")){
                final <- substr(PF_species[i,j], start=startFR + nchar("FRANCAIS"), stop=nchar(PF_species[i,j]))
                
              }else{
                convert_res[step,"Result"] <- emo::ji("prohibited")
                convert_res[step,"Message"] <- paste0("The summary was requested in English, but information in the '", colnames(PF_species)[j], "' field (3. SPECIES tab) is not provided in English. Please enter the information in English, preceded by the text 'ENGLISH -'.")
                KBAforms[step] <- NA
                error <- T
                break
              }
              
            }else if(checkEN){
              
              if(language == "english"){
                final <- substr(PF_species[i,j], start=startEN + nchar("ENGLISH"), stop=nchar(PF_species[i,j]))
                
              }else{
                convert_res[step,"Result"] <- emo::ji("prohibited")
                convert_res[step,"Message"] <- paste0("The summary was requested in French, but information in the '", colnames(PF_species)[j], "' field (3. SPECIES tab) is not provided in French. Please enter the information in French, preceded by the text 'FRANCAIS -'.")
                KBAforms[step] <- NA
                error <- T
                break
              }
            }
            
            # Trim front characters
            if(substr(final, start=1, stop=3) == " - "){
              final <- substr(final, start=4, stop=nchar(final))
            }
            
            if(substr(final, start=1, stop=2) == " -"){
              final <- substr(final, start=3, stop=nchar(final))
            }
            
            if(substr(final, start=1, stop=1) == "-"){
              final <- substr(final, start=2, stop=nchar(final))
            }
            
            # Trim white spaces
            final <- trimws(final, "both")
            
            # Assign to correct species entry
            PF_species[i,j] <- final
            
          }else{
            
            if(language=="french" & colnames(PF_species)[j] %in% c("Composition of 10 RUs", "Explanation of site estimates", "Explanation of reference estimates") & !is.na(PF_species[i,j])){
              convert_res[step,"Result"] <- emo::ji("prohibited")
              convert_res[step,"Message"] <- paste0("The summary was requested in French, but information in the '", colnames(PF_species)[j], "' field (3. SPECIES tab) is not provided in French. Please enter the information in French, preceded by the text 'FRANCAIS -'. Information in English should be preceded by 'ENGLISH -'.")
              KBAforms[step] <- NA
              error <- T
              break
            }
          }
        }
        
        if(error){
          break
        }
      }
      
      if(error){
        next
      }
      
                # Redact sensitive information
      PF_species %<>%
        mutate(Sensitive = F)
      
      if(reviewStage == "general"){
        
        # Check that the Public Display section is filled out
        if(sum(is.na(PF_species$display_taxonomicgroup), is.na(PF_species$display_taxonname), is.na(PF_species$display_assessmentinfo), is.na(PF_species$display_biodivelementdist)) > 0){
          convert_res[step,"Result"] <- emo::ji("prohibited")
          convert_res[step,"Message"] <- "You are requesting a summary for General Review and the Public Display section of the KBA Canada Proposal Form (SPECIES tab) is not filled out. Please fill out this section before you proceed with General Review."
          KBAforms[step] <- NA
          next
          
        }else{
          
          for(i in 1:nrow(PF_species)){
            
            alternativeName <- PF_species$display_alternativename[i] %>%
              str_to_sentence()
            if(language == "english"){
              alternativeName <- ifelse(is.na(alternativeName) || alternativeName == "", "A sensitive taxon", alternativeName)
            }else{
              alternativeName <- ifelse(is.na(alternativeName) || alternativeName == "", "Un taxon sensible", alternativeName)
            }
            
            # Display taxonomic group?
            if(PF_species$display_taxonomicgroup[i] == "No"){
              PF_species$`Taxonomic group`[i] <- "-"
              PF_species$`Common name`[i] <- alternativeName
              PF_species$`Scientific name`[i] <- alternativeName
              PF_species$Sensitive[i] <- T
            }
            
            # Display taxon name?
            if(PF_species$display_taxonname[i] == "No"){
              PF_species$`Common name`[i] <- alternativeName
              PF_species$`Scientific name`[i] <- alternativeName
              PF_species$Sensitive[i] <- T
            }
            
            # Display assessment information?
            if(PF_species$display_assessmentinfo[i] == "No"){
              PF_species$Status[i] <- "-"
              PF_species$`Status assessment agency`[i] <- "-"
              PF_species$`Reproductive Units (RU)`[i] <- "-"
              PF_species$`Assessment parameter`[i] <- "(i) -"
              PF_species$`Min site estimate`[i] <- "-"
              PF_species$`Best site estimate`[i] <- "-"
              PF_species$`Max site estimate`[i] <- "-"
              PF_species$`Year of site estimate`[i] <- "-"
              PF_species$`Min reference estimate`[i] <- "-"
              PF_species$`Best reference estimate`[i] <- "-"
              PF_species$`Max reference estimate`[i] <- "-"
              PF_species$`Composition of 10 RUs`[i] <- "-"
              PF_species$`RU source`[i] <- "-"
              PF_species$`Derivation of best estimate`[i] <- "-"
              PF_species$`Explanation of site estimates`[i] <- "-"
              PF_species$`Sources of site estimates`[i] <- "-"
              PF_species$`Explanation of reference estimates`[i] <- "-"
              PF_species$`Sources of reference estimates`[i] <- "-"
              PF_species$Sensitive[i] <- T
            }
          }
        }
      }
      
                # Sort by scientific name
      PF_species %<>% arrange(`Scientific name`)
    }
    
          # 4. ECOSYSTEMS & C
    if(nrow(PF_ecosystems) > 0){
      
                # If French is requested, translate the ecosystem type name to French
      if(language == "french"){
        
        PF_ecosystems %<>%
          left_join(., masterEcosystemList[,c("CNVC_ENGLISH_NAME", "CNVC_FRENCH_NAME")], by=c("Name of ecosystem type" = "CNVC_ENGLISH_NAME"))
        
        if(sum(is.na(PF_ecosystems$CNVC_FRENCH_NAME)) > 0){
          
          if(!sum(PF_ecosystems$`Name of ecosystem type` %in% masterEcosystemList$CNVC_ENGLISH_NAME) == nrow(PF_ecosystems)){
            convert_res[step,"Result"] <- emo::ji("prohibited")
            convert_res[step,"Message"] <- "Some values for the name of the ecosystem type (4. ECOSYSTEMS & C tab) are not recognized. Please cross-check your entries with the master ecosystems list."
            KBAforms[step] <- NA
            next
            
          }else{
            convert_res[step,"Result"] <- emo::ji("prohibited")
            convert_res[step,"Message"] <- "Some ecosystem types do not have a French common name in the master ecosystems list. Please contact Chloé and provide this error message."
            KBAforms[step] <- NA
            next
          }
        }
        
        PF_ecosystems %<>%
          mutate(`Name of ecosystem type` = CNVC_FRENCH_NAME) %>%
          select(-CNVC_FRENCH_NAME)
      }
      
                # If French is requested, translate the ecosystem classification level to French
      if(language == "french"){
        
        PF_ecosystems %<>%
          mutate(`Ecosystem level` = case_when(trimws(tolower(`Ecosystem level`)) == "group" ~ "Groupe",
                                               trimws(tolower(`Ecosystem level`)) == "alliance" ~ "Alliance",
                                               trimws(tolower(`Ecosystem level`)) == "ivc group" ~ "Groupe de la Classification internationale de la végétation (IVC)",
                                               trimws(tolower(`Ecosystem level`)) == "ivc alliance" ~ "Alliance de la Classification internationale de la végétation (IVC)",
                                               trimws(tolower(`Ecosystem level`)) == "cnvc group" ~ "Groupe de la Classification nationale de la végétation du Canada (CNVC)",
                                               trimws(tolower(`Ecosystem level`)) == "cnvc alliance" ~ "Alliance de la Classification nationale de la végétation du Canada (CNVC)",
                                               .default = 'Unrecognized'))
        
        if(nrow(PF_ecosystems %>% filter(`Ecosystem level` == "Unrecognized")) > 0){
          convert_res[step,"Result"] <- emo::ji("prohibited")
          convert_res[step,"Message"] <- "Some ecosystem level values (4. ECOSYSTEMS & C tab) are not recognized. Please enter one of 'Group', 'Alliance', 'IVC Group', 'IVC Alliance', 'CNVC Group', 'CNVC Alliance', or contact Chloé."
          KBAforms[step] <- NA
          next
        }
      }
      
                # Only retain information in the desired language
      error <- F
      
      for(i in 1:nrow(PF_ecosystems)){
        
        for(j in 1:ncol(PF_ecosystems)){
          
          if(grepl("FRANCAIS", PF_ecosystems[i,j]) | grepl("ENGLISH", PF_ecosystems[i,j])){
            
            # Initiate language check
            checkFR <- F
            checkEN <- F
            
            # Get index of FRANCAIS annotation
            if(grepl("FRANCAIS", PF_ecosystems[i,j])){
              checkFR <- T
              startFR <- unlist(gregexpr("FRANCAIS", PF_ecosystems[i,j]))
            }
            
            # Get index of ENGLISH annotation
            if(grepl("ENGLISH", PF_ecosystems[i,j])){
              checkEN <- T
              startEN <- unlist(gregexpr("ENGLISH", PF_ecosystems[i,j]))
            }
            
            # Get desired text
            if(checkFR & checkEN){
              
              if(startFR < startEN){
                FR <- substr(PF_ecosystems[i,j], start=startFR + nchar("FRANCAIS"), stop=startEN-1)
                EN <- substr(PF_ecosystems[i,j], start=startEN + nchar("ENGLISH"), stop=nchar(PF_ecosystems[i,j]))
                
              }else{
                FR <- substr(PF_ecosystems[i,j], start=startFR + nchar("FRANCAIS"), stop=nchar(PF_ecosystems[i,j]))
                EN <- substr(PF_ecosystems[i,j], start=startEN + nchar("ENGLISH"), stop=startFR-1)
              }
              
              if(language == "english"){
                final <- EN
                
              }else{
                final <- FR
              }
              
            }else if(checkFR){
              
              if(language == "french"){
                final <- substr(PF_ecosystems[i,j], start=startFR + nchar("FRANCAIS"), stop=nchar(PF_ecosystems[i,j]))
                
              }else{
                convert_res[step,"Result"] <- emo::ji("prohibited")
                convert_res[step,"Message"] <- paste0("The summary was requested in English, but information in the '", colnames(PF_ecosystems)[j], "' field (4. ECOSYSTEMS & C tab) is not provided in English. Please enter the information in English, preceded by the text 'ENGLISH -'.")
                KBAforms[step] <- NA
                error <- T
                break
              }
              
            }else if(checkEN){
              
              if(language == "english"){
                final <- substr(PF_ecosystems[i,j], start=startEN + nchar("ENGLISH"), stop=nchar(PF_ecosystems[i,j]))
                
              }else{
                convert_res[step,"Result"] <- emo::ji("prohibited")
                convert_res[step,"Message"] <- paste0("The summary was requested in French, but information in the '", colnames(PF_ecosystems)[j], "' field (4. ECOSYSTEMS & C tab) is not provided in French. Please enter the information in French, preceded by the text 'FRANCAIS -'.")
                KBAforms[step] <- NA
                error <- T
                break
              }
            }
            
            # Trim front characters
            if(substr(final, start=1, stop=3) == " - "){
              final <- substr(final, start=4, stop=nchar(final))
            }
            
            if(substr(final, start=1, stop=2) == " -"){
              final <- substr(final, start=3, stop=nchar(final))
            }
            
            if(substr(final, start=1, stop=1) == "-"){
              final <- substr(final, start=2, stop=nchar(final))
            }
            
            # Trim white spaces
            final <- trimws(final, "both")
            
            # Assign to correct species entry
            PF_ecosystems[i,j] <- final
            
          }else{
            
            if(language=="french" & colnames(PF_ecosystems)[j] %in% c("Ecosystem level justification") & !is.na(PF_ecosystems[i,j])){
              convert_res[step,"Result"] <- emo::ji("prohibited")
              convert_res[step,"Message"] <- paste0("The summary was requested in French, but information in the '", colnames(PF_ecosystems)[j], "' field (4. ECOSYSTEMS & C tab) is not provided in French. Please enter the information in French, preceded by the text 'FRANCAIS -'. Information in English should be preceded by 'ENGLISH -'.")
              KBAforms[step] <- NA
              error <- T
              break
            }
          }
        }
        
        if(error){
          break
        }
      }
      
      if(error){
        next
      }
      
                # Sort by ecosystem type name
      PF_ecosystems %<>% arrange(`Name of ecosystem type`)
    }
    
          # 7. CITATIONS
                # Redact sensitive citations
    if(reviewStage == "general"){
      PF_citations %<>%
        filter(!Sensitive == 1)
    }
    
    # Check which types of triggers exist
    triggerSpecies <- PF_species %>%
      filter(!is.na(`Criteria met`)) %>%
      pull(`Scientific name`) %>%
      unique() %>%
      length()
    
    triggerEcosystems <- PF_ecosystems %>%
      filter(!is.na(`Criteria met`)) %>%
      pull(`Name of ecosystem type`) %>%
      unique() %>%
      length()
    
    # Get trigger levels
          # Global species
    gS <- PF_species %>%
      filter(!is.na(`Criteria met`)) %>%
      pull(`Criteria met`) %>%
      unique() %>%
      grepl("g", ., fixed=T) %>%
      any()
    
          # National species
    nS <- PF_species %>%
      filter(!is.na(`Criteria met`)) %>%
      pull(`Criteria met`) %>%
      unique() %>%
      grepl("n", ., fixed=T) %>%
      any()
    
          # Global ecosystems
    gE <- PF_ecosystems %>%
      filter(!is.na(`Criteria met`)) %>%
      pull(`Criteria met`) %>%
      unique() %>%
      grepl("g", ., fixed=T) %>%
      any()
    
          # National ecosystems
    nE <- PF_ecosystems %>%
      filter(!is.na(`Criteria met`)) %>%
      pull(`Criteria met`) %>%
      unique() %>%
      grepl("n", ., fixed=T) %>%
      any()
    
    # Prepare variables
          # 1. KBA Name
    nationalName <<- PF_site$GENERAL[which(PF_site$Field == "National name")]
    
          # 2. Location
                # Jurisdiction
    juris <<- PF_site$GENERAL[which(PF_site$Field == "Province or Territory")]
    
    if(language == "french"){
      juris <<- xwalk_jurisdiction %>%
        .[which(.$Province_EN == juris), "Province_FR"]
    }
    
                # Latitude and Longitude
    lat <<- PF_site$GENERAL[which(PF_site$Field == "Latitude (dd.dddd)")] %>%
      as.numeric(.) %>%
      round(., 3)
    if(language == "english"){
      lat <<- ifelse(is.na(lat), "coordinates unspecified", lat)
    }else{
      lat <<- ifelse(is.na(lat), "coordonnées non spécifiées", lat)
    }
    
    lon <<- PF_site$GENERAL[which((PF_site$Field == "Longitude (dd.dddd)" | PF_site$Field == "Longitude (ddd.dddd)"))] %>%
      as.numeric(.) %>%
      round(., 3)
    lon <<- ifelse(is.na(lon), "", paste0("/", lon))
    
          # 3. KBA Scope
    criteriaMet <- PF_home$X4[which(PF_home$X3 == "Criteria met")]
    if(language == "english"){
      scope <<- ifelse(grepl("g", criteriaMet, fixed=T) & grepl("n", criteriaMet, fixed=T),
                       "Global and National",
                       ifelse(grepl("g", criteriaMet, fixed=T),
                              "Global",
                              "National"))
    }else{
      scope <<- ifelse(grepl("g", criteriaMet, fixed=T) & grepl("n", criteriaMet, fixed=T),
                       "Mondial et National",
                       ifelse(grepl("g", criteriaMet, fixed=T),
                              "Mondial",
                              "National"))
    }
    
          # 4. Proposal Development Lead
    proposalLead <<- PF_proposer$Entry[which(PF_proposer$Field == "Name of proposal development lead")]
    
    if(!is.na(PF_proposer$Entry[which(PF_proposer$Field == "Names and affiliations")])){
      proposalLead <<- trimws(proposalLead) %>%
        paste0(., ifelse(language=="english", ". Co-proposer(s): ", ". Co-développé avec : "), PF_proposer$Entry[which(PF_proposer$Field == "Names and affiliations")], ".")
    }
    
          # 7. Site Description
    if(language == "english"){
      siteDescription <<- PF_site$GENERAL[which(PF_site$Field == "Site description")]
    }else{
      siteDescription <<- PF_site$FRENCH[which(PF_site$Field == "Site description")]
    }
    
          # 8. Assessment Details - Level
    if(language == "english"){
      
      includeGlobalTriggers_eco <<- ifelse(gE, "GLOBAL", "")
      includeNationalTriggers_eco <<- ifelse(nE, "NATIONAL", "")
      includeGlobalTriggers_spp <<- ifelse(gS, "GLOBAL", "")
      includeNationalTriggers_spp <<- ifelse(nS, "NATIONAL", "")
      
    }else{
      
      includeGlobalTriggers_eco <<- ifelse(gE, "NIVEAU MONDIAL", "")
      includeNationalTriggers_eco <<- ifelse(nE, "NIVEAU NATIONAL", "")
      includeGlobalTriggers_spp <<- ifelse(gS, "NIVEAU MONDIAL", "")
      includeNationalTriggers_spp <<- ifelse(nS, "NIVEAU NATIONAL", "")
    }
    
          # 10. Delineation Rationale
    if(language == "english"){
      delineationRationale <<- PF_site$GENERAL[which(PF_site$Field == "Delineation rationale")]
    }else{
      delineationRationale <<- PF_site$FRENCH[which(PF_site$Field == "Delineation rationale")]
    }
    
          # 13. Additional Site Information
    if(language == "english"){
      nominationRationale <- PF_site$GENERAL[which(PF_site$Field == "Rationale for nomination")]
      additionalBiodiversity <- PF_site$GENERAL[which(PF_site$Field == "Additional biodiversity")]
      customaryJurisdiction <- PF_site$GENERAL[which(PF_site$Field == "Customary jurisdiction")]
    }else{
      nominationRationale <- PF_site$FRENCH[which(PF_site$Field == "Rationale for nomination")]
      additionalBiodiversity <- PF_site$FRENCH[which(PF_site$Field == "Additional biodiversity")]
      customaryJurisdiction <- PF_site$FRENCH[which(PF_site$Field == "Customary jurisdiction")]
    }
    
    if(language == "english"){
      customaryJurisdictionSource <- PF_site$GENERAL[which(PF_site$Field == "Customary jurisdiction source")]
      siteHistory <- PF_site$GENERAL[which(PF_site$Field == "Site history")]
      conservation <- PF_site$GENERAL[which(PF_site$Field == "Conservation")]
    }else{
      customaryJurisdictionSource <- PF_site$FRENCH[which(PF_site$Field == "Customary jurisdiction source")]
      siteHistory <- PF_site$FRENCH[which(PF_site$Field == "Site history")]
      conservation <- PF_site$FRENCH[which(PF_site$Field == "Conservation")]
    }
    
    # Prepare flextables
          # Site description
    siteDescription_ft <- siteDescription %>%
      as.data.frame() %>%
      flextable() %>%
      width(j=colnames(.), width=9) %>%
      delete_part(., part = "header") %>%
      font(fontname="Calibri", part="body") %>%
      fontsize(size=11, part='body') %>%
      border_remove()
    
          # Delineation rationale
    delineationRationale_ft <- delineationRationale %>%
      as.data.frame() %>%
      flextable() %>%
      width(j=colnames(.), width=9) %>%
      delete_part(., part = "header") %>%
      font(fontname="Calibri", part="body") %>%
      fontsize(size=11, part='body') %>%
      border_remove()
    
          # Criteria information
                # Check that at least one criterion is met
    if(is.na(criteriaMet)){
      convert_res[step,"Result"] <- emo::ji("prohibited")
      convert_res[step,"Message"] <- "No KBA Criteria met. Please revise your form and ensure that at least one criterion is met. If you believe that a KBA criterion should be met based on the information you provided in the form, contact Chloé and provide her with the error message."
      KBAforms[step] <- NA
      next
    }
    
                # Criteria definitions
    criteriaInfo <- data.frame(CriteriaFull = strsplit(criteriaMet, "; ")[[1]]) %>%
      mutate(Scope = ifelse(grepl("g", CriteriaFull, fixed=T), "Global", "National")) %>%
      mutate(Criteria = sapply(CriteriaFull, function(x) substr(x, start=2, stop=nchar(x)))) %>%
      arrange(Scope, Criteria) %>%
      mutate(Definition = sapply(1:nrow(.), function(x) xwalk_criterion[which(xwalk_criterion$Criterion == .$Criteria[x]), paste0(.$Scope[x], ifelse(language=="english", "_EN", "_FR"))]))
    
    if(language == "french"){
      criteriaInfo %<>%
        mutate(Scope = ifelse(Scope == "Global",
                              "Mondial",
                              Scope))
    }
    
                # Species criteria
    if(triggerSpecies > 0){
      
                      # Number of species
      maxColSpp <- max(sapply(PF_species$`Criteria met`[which(!is.na(PF_species$`Criteria met`))], function(x) str_count(x, ";")))+1
      criteriaColsSpp <- paste0("Col", 1:maxColSpp)
      
      criteriaInfoSpp <- PF_species %>%
        filter(!is.na(`Criteria met`)) %>%
        select(`Scientific name`, `Criteria met`) %>%
        separate(`Criteria met`, into=criteriaColsSpp, sep="; ", fill="right") %>%
        pivot_longer(all_of(criteriaColsSpp), names_to = "Remove", values_to="Criteria met") %>%
        filter(!is.na(`Criteria met`)) %>%
        group_by(`Criteria met`) %>%
        summarise(NTriggers = n(), .groups="drop") %>%
        left_join(criteriaInfo, ., by=c("CriteriaFull" = "Criteria met")) %>%
        filter(!is.na(NTriggers)) %>%
        mutate(Type = "Species")
      
                      # Species names
      criteriaInfoSpp_names <- PF_species %>%
        filter(!is.na(`Criteria met`)) %>%
        select(`Common name`, `Scientific name`, `Criteria met`) %>%
        separate(`Criteria met`, into=criteriaColsSpp, sep="; ", fill="right") %>%
        pivot_longer(all_of(criteriaColsSpp), names_to = "Remove", values_to="Criteria met") %>%
        filter(!is.na(`Criteria met`)) %>%
        arrange(`Criteria met`, `Scientific name`) %>%
        select(-Remove)
    }
    
                # Ecosystem criteria
    if(triggerEcosystems > 0){
      
                      # Number of ecosystems
      maxColEco <- max(sapply(PF_ecosystems$`Criteria met`[which(!is.na(PF_ecosystems$`Criteria met`))], function(x) str_count(x, ";")))+1
      criteriaColsEco <- paste0("Col", 1:maxColEco)
      
      criteriaInfoEco <- PF_ecosystems %>%
        filter(!is.na(`Criteria met`)) %>%
        select(`Name of ecosystem type`, `Criteria met`) %>%
        separate(`Criteria met`, into=criteriaColsEco, sep="; ", fill="right") %>%
        pivot_longer(all_of(criteriaColsEco), names_to = "Remove", values_to="Criteria met") %>%
        filter(!is.na(`Criteria met`)) %>%
        group_by(`Criteria met`) %>%
        summarise(NTriggers = n(), .groups="drop") %>%
        left_join(criteriaInfo, ., by=c("CriteriaFull" = "Criteria met")) %>%
        filter(!is.na(NTriggers))
      
                      # Ecosystem names
      criteriaInfoEco <- PF_ecosystems %>%
        filter(!is.na(`Criteria met`)) %>%
        select(`Name of ecosystem type`, `Criteria met`) %>%
        separate(`Criteria met`, into=criteriaColsEco, sep="; ", fill="right") %>%
        pivot_longer(all_of(criteriaColsEco), names_to = "Remove", values_to="Criteria met") %>%
        filter(!is.na(`Criteria met`)) %>%
        arrange(`Name of ecosystem type`) %>%
        group_by(`Criteria met`) %>%
        summarise(triggerNames = paste(`Name of ecosystem type`, collapse=", "), .groups="drop") %>%
        left_join(criteriaInfoEco, ., by=c("CriteriaFull" = "Criteria met")) %>%
        mutate(Type = "Ecosystem")
    }
    
                # All criteria
    if((triggerSpecies > 0) & (triggerEcosystems) > 0){
      criteriaInfo <- bind_rows(criteriaInfoSpp, criteriaInfoEco) %>%
        arrange(CriteriaFull)
      
    }else if(triggerSpecies > 0){
      criteriaInfo <- criteriaInfoSpp
      
    }else{
      criteriaInfo <- criteriaInfoEco
    }
    
                # Flextable
    criteriaInfo_ft <- criteriaInfo %>%
      mutate(Label = "") %>%
      mutate(Blank = "") %>%
      flextable(col_keys = c("Blank", "Label"))
    
    for(row in 1:nrow(criteriaInfo)){
      
      criterion <- criteriaInfo$CriteriaFull[row]
      type <- criteriaInfo$Type[row]
      
      if(type == "Species"){
        
        speciesNames <- criteriaInfoSpp_names %>%
          filter(`Criteria met` == criterion) %>%
          select(-`Criteria met`) %>%
          mutate(`Common name` = gsub("'", "\\'", `Common name`, fixed=T))
                 
        speciesNames_call <- ""
        
        for(sp in 1:nrow(speciesNames)){
          
          speciesNames_call <- paste0(speciesNames_call, ", as_chunk(x='", speciesNames$`Common name`[sp], " ('), as_chunk(x='", speciesNames$`Scientific name`[sp], "', props=fp_text(font.size=11, font.family='Calibri', italic=T)), as_chunk(x='), ')")
        }
        rm(sp)
        
        speciesNames_call <- substr(speciesNames_call, start=3, stop=nchar(speciesNames_call)-4) %>%
          paste0(., "')")
        
        if(language == "english"){
          compose_call <- paste0("criteriaInfo_ft %<>% compose(i=row, j='Label', value=as_paragraph(as_chunk(x=paste0(as.character('\u25CF'), ' ', criteriaInfo$Scope[row], ' ', criteriaInfo$Criteria[row], ' - ', criteriaInfo$Definition[row], ' [criterion met by ', criteriaInfo$NTriggers[row], ifelse(criteriaInfo$NTriggers[row] == 1, ' taxon: ', ' taxa: '))), ", speciesNames_call, ", as_chunk(x='].')))")
        }else{
          compose_call <- paste0("criteriaInfo_ft %<>% compose(i=row, j='Label', value=as_paragraph(as_chunk(x=paste0(as.character('\u25CF'), ' ', criteriaInfo$Scope[row], ' ', criteriaInfo$Criteria[row], ' - ', criteriaInfo$Definition[row], ' [critère rempli par ', criteriaInfo$NTriggers[row], ifelse(criteriaInfo$NTriggers[row] == 1, ' taxon : ', ' taxons : '))), ", speciesNames_call, ", as_chunk(x='].')))")
        }
        eval(parse(text=compose_call))
        
      }else{
        
        if(language == "english"){
          criteriaInfo_ft %<>% compose(i=row, j='Label', value=as_paragraph(as_chunk(x=paste0(as.character("\u25CF"), " ", criteriaInfo$Scope[row], " ", criteriaInfo$Criteria[row], " - ", criteriaInfo$Definition[row], " [criterion met by ", criteriaInfo$NTriggers[row], ifelse(criteriaInfo$NTriggers[row] == 1, " ecosystem type: ", " ecosystem types: "))), as_chunk(x=criteriaInfo$triggerNames[row], props=fp_text(font.size=11, font.family='Calibri', italic=F)), as_chunk(x="].")))
        }else{
          criteriaInfo_ft %<>% compose(i=row, j='Label', value=as_paragraph(as_chunk(x=paste0(as.character("\u25CF"), " ", criteriaInfo$Scope[row], " ", criteriaInfo$Criteria[row], " - ", criteriaInfo$Definition[row], " [critère rempli par ", criteriaInfo$NTriggers[row], ifelse(criteriaInfo$NTriggers[row] == 1, " type d'écosystème : ", " types d'écosystèmes : "))), as_chunk(x=criteriaInfo$triggerNames[row], props=fp_text(font.size=11, font.family='Calibri', italic=F)), as_chunk(x="].")))
        }
      }
    }
    
    criteriaInfo_ft %<>% 
      font(fontname="Calibri", part="body") %>%
      fontsize(size=11, part='body') %>%
      width(j=colnames(.), width=c(0.3, 9)) %>%
      delete_part(part='header') %>%
      border_remove() %>%
      align(j=2, align = "left", part = "body")
    
          # Species assessments
    if(triggerSpecies > 0){
      
                # Get information
      speciesAssessments <- PF_species %>%
        filter(!is.na(`Criteria met`)) %>%
        mutate_at(vars(`Reproductive Units (RU)`, `Min site estimate`, `Best site estimate`, `Max site estimate`, `Min reference estimate`, `Best reference estimate`, `Max reference estimate`), as.double) %>%
        mutate(PercentAtSite = ifelse(`Best site estimate` == "-", "-", round(100 * `Best site estimate`/`Best reference estimate`, 1))) %>%
        mutate(Blank = "") %>%
        mutate(Status = ifelse(grepl("A1", `Criteria met`, fixed=T),
                               ifelse(`Status assessment agency` == "-", "-", paste0(Status, " (", `Status assessment agency`, ")")),
                               ifelse(language == "english",
                                      "Not applicable",
                                      "Non applicable"))) %>%
        mutate(SiteEstimate_Min = as.character(`Min site estimate`),
               SiteEstimate_Best = as.character(`Best site estimate`),
               SiteEstimate_Max = as.character(`Max site estimate`),
               TotalEstimate_Min = as.character(`Min reference estimate`),
               TotalEstimate_Best = as.character(`Best reference estimate`),
               TotalEstimate_Max = as.character(`Max reference estimate`)) %>%
        mutate(AssessmentParameter = sapply(`Assessment parameter`, function(x) str_to_sentence(substr(x, start=str_locate(x, "\\)")[1,1]+2, stop=nchar(x)))))
      
      if(language == "french"){
        speciesAssessments %<>%
          left_join(., xwalk_assessmentParameter[,c("AssessmentParameter_EN", "AssessmentParameter_FR")], by=c("AssessmentParameter" = "AssessmentParameter_EN")) %>%
          mutate(AssessmentParameter = AssessmentParameter_FR) %>%
          select(-AssessmentParameter_FR)
      }
      
      speciesAssessments %<>%
        mutate(AssessmentParameter = ifelse(AssessmentParameter %in% c("Area of occupancy", "Zone d'occupation", "Extent of suitable habitat", "Étendue de l'habitat approprié", "Range", "Aire de répartition"), paste(AssessmentParameter, "(km2)"), AssessmentParameter)) %>%
        mutate(`Name` = paste0(`Common name`, " (", `Scientific name`, ")")) %>%
        select(Name, Status, `Criteria met`, `Description of evidence`, `Reproductive Units (RU)`, `Composition of 10 RUs`, `RU source`, AssessmentParameter, Blank, SiteEstimate_Min, SiteEstimate_Best, SiteEstimate_Max, `Year of site estimate`, `Derivation of best estimate`, `Explanation of site estimates`, `Sources of site estimates`, TotalEstimate_Min, TotalEstimate_Best, TotalEstimate_Max, `Explanation of reference estimates`, `Sources of reference estimates`, PercentAtSite, Sensitive)
      
                # Separate global and national assessments
      speciesAssessments_g <- speciesAssessments %>%
        filter(grepl("g", `Criteria met`, fixed=T)) %>%
        mutate(`Criteria met` = gsub("g", "", `Criteria met`))
      
      speciesAssessments_n <- speciesAssessments %>%
        filter(grepl("n", `Criteria met`, fixed=T)) %>%
        mutate(`Criteria met` = gsub("n", "", `Criteria met`))
      
      if(!(nrow(speciesAssessments_g) + nrow(speciesAssessments_n)) == nrow(speciesAssessments)){
        convert_res[step,"Result"] <- emo::ji("prohibited")
        convert_res[step,"Message"] <- "Some species assessments are not being correctly classified as global or national assessments. This is an error with the code. Please contact Chloé and provide her with this error message."
        KBAforms[step] <- NA
        next
      }
      rm(speciesAssessments)
      
                # Information for the footnotes
      footnotesSpecies_g <- speciesAssessments_g %>%
        select(`Composition of 10 RUs`, `Description of evidence`, `RU source`, `Derivation of best estimate`, `Explanation of site estimates`, `Sources of site estimates`, `Explanation of reference estimates`, `Sources of reference estimates`, Sensitive, `Criteria met`) %>%
        mutate(`Composition of 10 RUs` = sapply(`Composition of 10 RUs`, function(x) ifelse(substr(x, start=nchar(x), stop=nchar(x)) == ".", x, paste0(x, ".")))) %>%
        mutate(`Description of evidence` = sapply(`Description of evidence`, function(x) ifelse(substr(x, start=nchar(x), stop=nchar(x)) == ".", x, paste0(x, ".")))) %>%
        mutate(`RU source` = sapply(`RU source`, function(x) ifelse(substr(x, start=nchar(x), stop=nchar(x)) == ".", x, paste0(x, ".")))) %>%
        mutate(`Derivation of best estimate` = sapply(`Derivation of best estimate`, function(x) ifelse(substr(x, start=nchar(x), stop=nchar(x)) == ".", x, paste0(x, ".")))) %>% 
        mutate(`Explanation of site estimates` = sapply(`Explanation of site estimates`, function(x) ifelse(substr(x, start=nchar(x), stop=nchar(x)) == ".", x, paste0(x, ".")))) %>%
        mutate(`Sources of site estimates` = sapply(`Sources of site estimates`, function(x) ifelse(substr(x, start=nchar(x), stop=nchar(x)) == ".", x, paste0(x, ".")))) %>%
        mutate(`Explanation of reference estimates` = sapply(`Explanation of reference estimates`, function(x) ifelse(substr(x, start=nchar(x), stop=nchar(x)) == ".", x, paste0(x, ".")))) %>%
        mutate(`Sources of reference estimates` = sapply(`Sources of reference estimates`, function(x) ifelse(substr(x, start=nchar(x), stop=nchar(x)) == ".", x, paste0(x, "."))))
      
      if(language == "english"){
        footnotesSpecies_g %<>%
          mutate(RU_Source = paste0(ifelse(!is.na(`Composition of 10 RUs`), paste0("Composition of 10 Reproductive Units (RUs): ", `Composition of 10 RUs`), ""), ifelse(!is.na(`Description of evidence`), paste0(" Description of evidence: ", `Description of evidence`), ""), ifelse(!is.na(`RU source`), paste0(" Source of RU data: ", `RU source`), ""))) %>%
          mutate(RU_Source = ifelse(RU_Source == "", NA, RU_Source)) %>%
          mutate(RU_Source = trimws(RU_Source)) %>%
          mutate(Site_Source = paste0(ifelse(!is.na(`Derivation of best estimate`), paste0("Derivation of site estimate: ", `Derivation of best estimate`), ""), ifelse(!is.na(`Explanation of site estimates`), paste0(" Explanation of site estimate(s): ", `Explanation of site estimates`), ""), ifelse(!is.na(`Sources of site estimates`), paste0(" Source(s) of site estimate(s): ", `Sources of site estimates`), ""))) %>%
          mutate(Site_Source = ifelse(Site_Source == "", NA, Site_Source)) %>%
          mutate(Reference_Source = paste0(ifelse(!is.na(`Explanation of reference estimates`), paste0("Explanation of global estimate(s): ", `Explanation of reference estimates`), ""), ifelse(!is.na(`Sources of reference estimates`), paste0(" Source(s) of global estimate(s): ", `Sources of reference estimates`), ""))) %>%
          mutate(Reference_Source = ifelse(Reference_Source == "", NA, Reference_Source)) %>%
          mutate(D1b = ifelse(grepl("D1b", `Criteria met`, fixed=T), "Meets criterion D1b because it is one of 10 largest aggregations in the world for this species.", NA)) %>%
          mutate(Footnote = paste0(ifelse(!is.na(D1b), paste0("CRITERIA MET\n", D1b, "\n\n"), ""), ifelse(!is.na(RU_Source), paste0("REPRODUCTIVE UNITS\n", trimws(RU_Source), "\n\n"), ""), ifelse(!is.na(Site_Source), paste0("SITE ESTIMATE\n", trimws(Site_Source), "\n\n"), ""), ifelse(!is.na(Reference_Source), paste0("GLOBAL ESTIMATE\n", trimws(Reference_Source)), ""))) %>%
          mutate(Footnote = ifelse(Sensitive, "This species is considered sensitive. For more information, please contact the KBA Canada Secretariat.", Footnote)) %>%
          pull(Footnote)
        
      }else{
        footnotesSpecies_g %<>%
          mutate(RU_Source = paste0(ifelse(!is.na(`Composition of 10 RUs`), paste0("Composition de 10 Unités Reproductives (URs) : ", `Composition of 10 RUs`), ""), ifelse(!is.na(`Description of evidence`), paste0(" Description des éléments de preuve : ", `Description of evidence`), ""), ifelse(!is.na(`RU source`), paste0(" Source des données d'URs : ", `RU source`), ""))) %>%
          mutate(RU_Source = ifelse(RU_Source == "", NA, RU_Source)) %>%
          mutate(Site_Source = paste0(ifelse(!is.na(`Derivation of best estimate`), paste0("Calcul de l'estimation au site : ", `Derivation of best estimate`), ""), ifelse(!is.na(`Explanation of site estimates`), paste0(" Explication de(s) estimation(s) au site : ", `Explanation of site estimates`), ""), ifelse(!is.na(`Sources of site estimates`), paste0(" Source(s) de(s) estimation(s) au site : ", `Sources of site estimates`), ""))) %>%
          mutate(Site_Source = ifelse(Site_Source == "", NA, Site_Source)) %>%
          mutate(Reference_Source = paste0(ifelse(!is.na(`Explanation of reference estimates`), paste0("Explication de(s) estimation(s) mondiale(s) : ", `Explanation of reference estimates`), ""), ifelse(!is.na(`Sources of reference estimates`), paste0(" Source(s) de(s) estimation(s) mondiale(s) : ", `Sources of reference estimates`), ""))) %>%
          mutate(Reference_Source = ifelse(Reference_Source == "", NA, Reference_Source)) %>%
          mutate(D1b = ifelse(grepl("D1b", `Criteria met`, fixed=T), "Le site répond au critère D1b, parce qu'il contient une des 10 plus grandes aggrégations au monde pour cette espèce.", NA)) %>%
          mutate(Footnote = paste0(ifelse(!is.na(D1b), paste0("CRITÈRE ATTEINT\n", D1b, "\n\n"), ""), ifelse(!is.na(RU_Source), paste0("UNITÉS REPRODUCTRICES\n", trimws(RU_Source), "\n\n"), ""), ifelse(!is.na(Site_Source), paste0("ESTIMATION AU SITE\n", trimws(Site_Source), "\n\n"), ""), ifelse(!is.na(Reference_Source), paste0("ESTIMATION MONDIALE\n", trimws(Reference_Source)), ""))) %>%
          mutate(Footnote = ifelse(Sensitive, "Cette espèce est considérée comme une espèce sensible. Pour d'avantage d'informations, merci de contacter le Secrétariat KBA Canada.", Footnote)) %>%
          pull(Footnote)
      }
      
      footnotesSpecies_n <- speciesAssessments_n %>%
        select(`Composition of 10 RUs`, `Description of evidence`, `RU source`, `Derivation of best estimate`, `Explanation of site estimates`, `Sources of site estimates`, `Explanation of reference estimates`, `Sources of reference estimates`, Sensitive, `Criteria met`) %>%
        mutate(`Composition of 10 RUs` = sapply(`Composition of 10 RUs`, function(x) ifelse(substr(x, start=nchar(x), stop=nchar(x)) == ".", x, paste0(x, ".")))) %>%
        mutate(`Description of evidence` = sapply(`Description of evidence`, function(x) ifelse(substr(x, start=nchar(x), stop=nchar(x)) == ".", x, paste0(x, ".")))) %>%
        mutate(`RU source` = sapply(`RU source`, function(x) ifelse(substr(x, start=nchar(x), stop=nchar(x)) == ".", x, paste0(x, ".")))) %>%
        mutate(`Derivation of best estimate` = sapply(`Derivation of best estimate`, function(x) ifelse(substr(x, start=nchar(x), stop=nchar(x)) == ".", x, paste0(x, ".")))) %>% 
        mutate(`Explanation of site estimates` = sapply(`Explanation of site estimates`, function(x) ifelse(substr(x, start=nchar(x), stop=nchar(x)) == ".", x, paste0(x, ".")))) %>%
        mutate(`Sources of site estimates` = sapply(`Sources of site estimates`, function(x) ifelse(substr(x, start=nchar(x), stop=nchar(x)) == ".", x, paste0(x, ".")))) %>%
        mutate(`Explanation of reference estimates` = sapply(`Explanation of reference estimates`, function(x) ifelse(substr(x, start=nchar(x), stop=nchar(x)) == ".", x, paste0(x, ".")))) %>%
        mutate(`Sources of reference estimates` = sapply(`Sources of reference estimates`, function(x) ifelse(substr(x, start=nchar(x), stop=nchar(x)) == ".", x, paste0(x, "."))))
      
      if(language == "english"){
        footnotesSpecies_n %<>%
          mutate(RU_Source = paste0(ifelse(!is.na(`Composition of 10 RUs`), paste0("Composition of 10 Reproductive Units (RUs): ", `Composition of 10 RUs`), ""), ifelse(!is.na(`Description of evidence`), paste0(" Description of evidence: ", `Description of evidence`), ""), ifelse(!is.na(`RU source`), paste0(" Source of RU data: ", `RU source`), ""))) %>%
          mutate(RU_Source = ifelse(RU_Source == "", NA, RU_Source)) %>%
          mutate(Site_Source = paste0(ifelse(!is.na(`Derivation of best estimate`), paste0("Derivation of site estimate: ", `Derivation of best estimate`), ""), ifelse(!is.na(`Explanation of site estimates`), paste0(" Explanation of site estimate(s): ", `Explanation of site estimates`), ""), ifelse(!is.na(`Sources of site estimates`), paste0(" Source(s) of site estimate(s): ", `Sources of site estimates`), ""))) %>%
          mutate(Site_Source = ifelse(Site_Source == "", NA, Site_Source)) %>%
          mutate(Reference_Source = paste0(ifelse(!is.na(`Explanation of reference estimates`), paste0("Explanation of national estimate(s): ", `Explanation of reference estimates`), ""), ifelse(!is.na(`Sources of reference estimates`), paste0(" Source(s) of national estimate(s): ", `Sources of reference estimates`), ""))) %>%
          mutate(Reference_Source = ifelse(Reference_Source == "", NA, Reference_Source)) %>%
          mutate(D1b = ifelse(grepl("D1b", `Criteria met`, fixed=T), "Meets criterion D1b because it is one of 10 largest aggregations in Canada for this taxon.", NA)) %>%
          mutate(Footnote = paste0(ifelse(!is.na(D1b), paste0("CRITERIA MET\n", D1b, "\n\n"), ""), ifelse(!is.na(RU_Source), paste0("REPRODUCTIVE UNITS\n", trimws(RU_Source), "\n\n"), ""), ifelse(!is.na(Site_Source), paste0("SITE ESTIMATE\n", trimws(Site_Source), "\n\n"), ""), ifelse(!is.na(Reference_Source), paste0("NATIONAL ESTIMATE\n", trimws(Reference_Source)), ""))) %>%
          mutate(Footnote = ifelse(Sensitive, "This species is considered sensitive. For more information, please contact the KBA Canada Secretariat.", Footnote)) %>%
          pull(Footnote)
        
      }else{
        footnotesSpecies_n %<>%
          mutate(RU_Source = paste0(ifelse(!is.na(`Composition of 10 RUs`), paste0("Composition de 10 Unités Reproductives (URs) : ", `Composition of 10 RUs`), ""), ifelse(!is.na(`Description of evidence`), paste0(" Description des éléments de preuve : ", `Description of evidence`), ""), ifelse(!is.na(`RU source`), paste0(" Source des données d'URs : ", `RU source`), ""))) %>%
          mutate(RU_Source = ifelse(RU_Source == "", NA, RU_Source)) %>%
          mutate(Site_Source = paste0(ifelse(!is.na(`Derivation of best estimate`), paste0("Calcul de l'estimation au site : ", `Derivation of best estimate`), ""), ifelse(!is.na(`Explanation of site estimates`), paste0(" Explication de(s) estimation(s) au site : ", `Explanation of site estimates`), ""), ifelse(!is.na(`Sources of site estimates`), paste0(" Source(s) de(s) estimation(s) au site : ", `Sources of site estimates`), ""))) %>%
          mutate(Site_Source = ifelse(Site_Source == "", NA, Site_Source)) %>%
          mutate(Reference_Source = paste0(ifelse(!is.na(`Explanation of reference estimates`), paste0("Explication de(s) estimation(s) nationale(s) : ", `Explanation of reference estimates`), ""), ifelse(!is.na(`Sources of reference estimates`), paste0(" Source(s) de(s) estimation(s) nationale(s) : ", `Sources of reference estimates`), ""))) %>%
          mutate(Reference_Source = ifelse(Reference_Source == "", NA, Reference_Source)) %>%
          mutate(D1b = ifelse(grepl("D1b", `Criteria met`, fixed=T), "Le site répond au critère D1b, parce qu'il contient une des 10 plus grandes aggrégations au Canada pour ce taxon.", NA)) %>%
          mutate(Footnote = paste0(ifelse(!is.na(D1b), paste0("CRITÈRE ATTEINT\n", D1b, "\n\n"), ""), ifelse(!is.na(RU_Source), paste0("UNITÉS REPRODUCTRICES\n", trimws(RU_Source), "\n\n"), ""), ifelse(!is.na(Site_Source), paste0("ESTIMATION AU SITE\n", trimws(Site_Source), "\n\n"), ""), ifelse(!is.na(Reference_Source), paste0("ESTIMATION NATIONALE\n", trimws(Reference_Source)), ""))) %>%
          mutate(Footnote = ifelse(Sensitive, "Cette espèce est considérée comme une espèce sensible. Pour d'avantage d'informations, merci de contacter le Secrétariat KBA Canada.", Footnote)) %>%
          pull(Footnote)
      }
      
                # Information for the main table
      speciesAssessments_g %<>% select(-c(`Composition of 10 RUs`, `Description of evidence`, `RU source`, `Derivation of best estimate`, `Explanation of site estimates`, `Sources of site estimates`, `Explanation of reference estimates`, `Sources of reference estimates`))
      
      speciesAssessments_n %<>% select(-c(`Composition of 10 RUs`, `Description of evidence`, `RU source`, `Derivation of best estimate`, `Explanation of site estimates`, `Sources of site estimates`, `Explanation of reference estimates`, `Sources of reference estimates`))
      
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
        
        speciesAssessments_g %<>%
          mutate(row = seq(1, nrow(.)*2, 2)) %>%
          add_row(Name = footnotesSpecies_g,
                  row = seq(2, nrow(.)*2, 2)) %>%
          arrange(row) %>%
          select(-row)
        
        if(bestOnly_g){
          speciesAssessments_g_ft <- speciesAssessments_g %>%
            select(-Sensitive) %>%
            flextable()
          
          if(language == "english"){
            speciesAssessments_g_ft %<>%
              width(j=colnames(.), width=c(1.5,1.3,0.65,1.3,1.3,0.05,0.7,0.7,0.7,0.8)) %>%
              set_header_labels(values=list(`Name` = "Species", Status = "Status*", `Criteria met`="Criteria Met", `Reproductive Units (RU)` = "# of Reproductive Units", AssessmentParameter = 'Assessment Parameter', Blank='', SiteEstimate_Best = "Value", `Year of site estimate` = "Year", TotalEstimate_Best = 'Global Estimate', PercentAtSite = "% of Global Pop. at Site")) %>%
              add_header_row(values = c("Species", "Status*", "Criteria Met", "# of Reproductive Units", "Assessment Parameter", "", "Site Estimate", 'Global Estimate', "% of Global Pop. at Site"), colwidths=c(1, 1, 1, 1, 1, 1, 2, 1, 1))
          }else{
            speciesAssessments_g_ft %<>%
              width(j=colnames(.), width=c(1.3,1.1,0.85,1.3,1.3,0.05,0.7,0.7,0.9,0.8)) %>%
              set_header_labels(values=list(`Name` = "Espèce", Status = "Statut*", `Criteria met`="Critère(s) atteint(s)", `Reproductive Units (RU)` = "# d’Unités Reproductives", AssessmentParameter = 'Paramètre d’évaluation', Blank='', SiteEstimate_Best = "Valeur", `Year of site estimate` = "Année", TotalEstimate_Best = 'Estimation mondiale', PercentAtSite = "% de la pop. mondiale au site")) %>%
              add_header_row(values = c("Espèce", "Statut*", "Critère(s) atteint(s)", "# d’Unités Reproductives", "Paramètre d’évaluation", "", "Estimation au site", "Estimation mondiale", "% de la pop. mondiale au site"), colwidths=c(1, 1, 1, 1, 1, 1, 2, 1, 1))
          }
          
          speciesAssessments_g_ft %<>%  
            align(align = "center", part="header") %>%
            font(fontname="Calibri", part="header") %>%
            fontsize(size=11, part='header') %>%
            bold(i=1, bold=T, part='header') %>%
            merge_v(part = "header") %>%
            font(fontname="Calibri", part="body") %>%
            fontsize(size=11, part='body') %>%
            hline_top(part="all") %>%
            border_remove() %>%
            hline(border = fp_border(width = 1), part="header") %>%
            hline_top(border = fp_border(width = 2), part="header") %>%
            hline_bottom(border = fp_border(width = 2), part="header") %>%
            hline_bottom(border=fp_border(width=1), part='body') %>%
            align(j=c(2,3,4,5,6,7,8,9,10), align = "center", part = "body")
          
          for(row in seq(2, nrow(speciesAssessments_g), 2)){
            
            speciesAssessments_g_ft %<>%
              merge_at(i=row, j=1:11) %>%
              bg(i=row, bg = "#EFEFEF", part = "body") %>%
              bold(i=row-1, j=1, bold=T, part="body")
          }
          
        }else{
          speciesAssessments_g_ft <- speciesAssessments_g %>%
            select(-Sensitive) %>%
            mutate(Blank2 = "") %>%
            relocate(Blank2, .after = `Year of site estimate`) %>%
            flextable()
          
          if(language == "english"){
            speciesAssessments_g_ft %<>%
              width(j=colnames(.), width=c(1.05,0.8,0.65,0.9,0.9,0.05,0.55,0.55,0.55,0.55,0.05,0.55,0.55,0.55,0.8)) %>%
              set_header_labels(values=list(`Name` = "Species", Status = "Status*", `Criteria met`="Criteria Met", `Reproductive Units (RU)` = "# of Reproductive Units", AssessmentParameter = 'Assessment Parameter', Blank='', SiteEstimate_Min = "Min", SiteEstimate_Best = "Best", SiteEstimate_Max = "Max", SiteEstimate_Year = "Year", Blank2 = "", TotalEstimate_Min = "Min", TotalEstimate_Best = "Best", TotalEstimate_Max = "Max", PercentAtSite = "% of Global Pop. at Site")) %>%
              add_header_row(values = c("Species", "Status*", "Criteria Met", "# of Reproductive Units", "Assessment Parameter", "", "Site Estimate", "", "Global Estimate", "% of Global Pop. at Site"), colwidths=c(1, 1, 1, 1, 1, 1, 4, 1, 3, 1))
          }else{
            speciesAssessments_g_ft %<>%
              width(j=colnames(.), width=c(1.1,0.8,0.85,1.1,0.95,0.05,0.45,0.5,0.45,0.55,0.05,0.45,0.5,0.45,0.8)) %>%
              set_header_labels(values=list(`Name` = "Espèce", Status = "Statut*", `Criteria met`="Critère(s) atteint(s)", `Reproductive Units (RU)` = "# d’Unités Reprod.", AssessmentParameter = 'Paramètre d’évaluation', Blank='', SiteEstimate_Min = "Min", SiteEstimate_Best = "Meilleure", SiteEstimate_Max = "Max", `Year of site estimate` = "Année", TotalEstimate_Min = "Min", TotalEstimate_Best = 'Meilleure', TotalEstimate_Max = "Max", PercentAtSite = "% de la pop. mondiale au site")) %>%
              add_header_row(values = c("Espèce", "Statut*", "Critère(s) atteint(s)", "# d’Unités Reprod.", "Paramètre d’évaluation", "", "Estimation au site", "", "Estimation mondiale", "% de la pop. mondiale au site"), colwidths=c(1, 1, 1, 1, 1, 1, 4, 1, 3, 1))
          }
          
          speciesAssessments_g_ft %<>%  
            align(align = "center", part="header") %>%
            font(fontname="Calibri", part="header") %>%
            fontsize(size=11, part='header') %>%
            bold(i=1, bold=T, part='header') %>%
            merge_v(part = "header") %>%
            font(fontname="Calibri", part="body") %>%
            fontsize(size=11, part='body') %>%
            hline_top(part="all") %>%
            border_remove() %>%
            hline(border = fp_border(width = 1), part="header") %>%
            hline_top(border = fp_border(width = 2), part="header") %>%
            hline_bottom(border = fp_border(width = 2), part="header") %>%
            hline_bottom(border=fp_border(width=1), part='body') %>%
            align(j=c(2,3,4,7,8,9,10,12,13,14,15), align = "center", part = "body")
          
          for(row in seq(2, nrow(speciesAssessments_g), 2)){
            
            speciesAssessments_g_ft %<>%
              merge_at(i=row, j=1:15) %>%
              bg(i=row, bg = "#EFEFEF", part = "body") %>%
              bold(i=row-1, j=1, bold=T, part="body")
          }
        }
        
      }else{
        speciesAssessments_g_ft <- ""
      }
      
                       # Species assessment - National
      if(nrow(speciesAssessments_n) > 0){
        
        speciesAssessments_n %<>%
          mutate(row = seq(1, nrow(.)*2, 2)) %>%
          add_row(Name = footnotesSpecies_n,
                  row = seq(2, nrow(.)*2, 2)) %>%
          arrange(row) %>%
          select(-row)
        
        if(bestOnly_n){
          speciesAssessments_n_ft <- speciesAssessments_n %>%
            select(-Sensitive) %>%
            flextable()
          
          if(language == "english"){
            speciesAssessments_n_ft %<>%
              width(j=colnames(.), width=c(1.5,1.3,0.65,1.3,1.3,0.05,0.7,0.7,0.7,0.8)) %>%
              set_header_labels(values=list(`Name` = "Taxon", Status = "Status*", `Criteria met`="Criteria Met", `Reproductive Units (RU)` = "# of Reproductive Units", AssessmentParameter = 'Assessment Parameter', Blank='', SiteEstimate_Best = "Value", `Year of site estimate` = "Year", TotalEstimate_Best = 'National Estimate', PercentAtSite = "% of National Pop. at Site")) %>%
              add_header_row(values = c("Taxon", "Status*", "Criteria Met", "# of Reproductive Units", "Assessment Parameter", "", "Site Estimate", 'National Estimate', "% of National Pop. at Site"), colwidths=c(1, 1, 1, 1, 1, 1, 2, 1, 1))
          }else{
            speciesAssessments_n_ft %<>%
              width(j=colnames(.), width=c(1.3,1.1,0.85,1.3,1.3,0.05,0.7,0.7,0.9,0.8)) %>%
              set_header_labels(values=list(`Name` = "Taxon", Status = "Statut*", `Criteria met`="Critère(s) atteint(s)", `Reproductive Units (RU)` = "# d’Unités Reproductives", AssessmentParameter = 'Paramètre d’évaluation', Blank='', SiteEstimate_Best = "Valeur", `Year of site estimate` = "Année", TotalEstimate_Best = 'Estimation nationale', PercentAtSite = "% de la pop. nationale au site")) %>%
              add_header_row(values = c("Taxon", "Statut*", "Critère(s) atteint(s)", "# d’Unités Reproductives", "Paramètre d’évaluation", "", "Estimation au site", "Estimation nationale", "% de la pop. nationale au site"), colwidths=c(1, 1, 1, 1, 1, 1, 2, 1, 1))
          }
          
          speciesAssessments_n_ft %<>%
            align(align = "center", part="header") %>%
            font(fontname="Calibri", part="header") %>%
            fontsize(size=11, part='header') %>%
            bold(i=1, bold=T, part='header') %>%
            merge_v(part = "header") %>%
            font(fontname="Calibri", part="body") %>%
            fontsize(size=11, part='body') %>%
            hline_top(part="all") %>%
            border_remove() %>%
            hline(border = fp_border(width = 1), part="header") %>%
            hline_top(border = fp_border(width = 2), part="header") %>%
            hline_bottom(border = fp_border(width = 2), part="header") %>%
            hline_bottom(border=fp_border(width=1), part='body') %>%
            align(j=c(2,3,4,5,6,7,8,9,10), align = "center", part = "body")
          
          for(row in seq(2, nrow(speciesAssessments_n), 2)){
            
            speciesAssessments_n_ft %<>%
              merge_at(i=row, j=1:11) %>%
              bg(i=row, bg = "#EFEFEF", part = "body") %>%
              bold(i=row-1, j=1, bold=T, part="body")
          }
          
        }else{
          speciesAssessments_n_ft <- speciesAssessments_n %>%
            select(-Sensitive) %>%
            mutate(Blank2 = "") %>%
            relocate(Blank2, .after = `Year of site estimate`) %>%
            flextable()
          
          if(language == "english"){
            speciesAssessments_n_ft %<>%
              width(j=colnames(.), width=c(1.05,0.8,0.65,0.9,0.9,0.05,0.55,0.55,0.55,0.55,0.05,0.55,0.55,0.55,0.8)) %>%
              set_header_labels(values=list(`Name` = "Taxon", Status = "Status*", `Criteria met`="Criteria Met", `Reproductive Units (RU)` = "# of Reproductive Units", AssessmentParameter = 'Assessment Parameter', Blank='', SiteEstimate_Min = "Min", SiteEstimate_Best = "Best", SiteEstimate_Max = "Max", SiteEstimate_Year = "Year", Blank2 = "", TotalEstimate_Min = "Min", TotalEstimate_Best = "Best", TotalEstimate_Max = "Max", PercentAtSite = "% of National Pop. at Site")) %>%
              add_header_row(values = c("Taxon", "Status*", "Criteria Met", "# of Reproductive Units", "Assessment Parameter", "", "Site Estimate", "", "National Estimate", "% of National Pop. at Site"), colwidths=c(1, 1, 1, 1, 1, 1, 4, 1, 3, 1))
          }else{
            speciesAssessments_n_ft %<>%
              width(j=colnames(.), width=c(1.1,0.8,0.85,1.1,0.95,0.05,0.45,0.5,0.45,0.55,0.05,0.45,0.5,0.45,0.8)) %>%
              set_header_labels(values=list(`Name` = "Taxon", Status = "Statut*", `Criteria met`="Critère(s) atteint(s)", `Reproductive Units (RU)` = "# d’Unités Reprod.", AssessmentParameter = 'Paramètre d’évaluation', Blank='', SiteEstimate_Min = "Min", SiteEstimate_Best = "Meilleure", SiteEstimate_Max = "Max", `Year of site estimate` = "Année", Blank2 = "", TotalEstimate_Min = "Min", TotalEstimate_Best = 'Meilleure', TotalEstimate_Max = "Max", PercentAtSite = "% de la pop. nationale au site")) %>%
              add_header_row(values = c("Taxon", "Statut*", "Critère(s) atteint(s)", "# d’Unités Reprod.", "Paramètre d’évaluation", "", "Estimation au site", "", "Estimation nationale", "% de la pop. nationale au site"), colwidths=c(1, 1, 1, 1, 1, 1, 4, 1, 3, 1))
          }
          
          speciesAssessments_n_ft %<>%
            align(align = "center", part="header") %>%
            font(fontname="Calibri", part="header") %>%
            fontsize(size=11, part='header') %>%
            bold(i=1, bold=T, part='header') %>%
            merge_v(part = "header") %>%
            font(fontname="Calibri", part="body") %>%
            fontsize(size=11, part='body') %>%
            hline_top(part="all") %>%
            border_remove() %>%
            hline(border = fp_border(width = 1), part="header") %>%
            hline_top(border = fp_border(width = 2), part="header") %>%
            hline_bottom(border = fp_border(width = 2), part="header") %>%
            hline_bottom(border=fp_border(width=1), part='body') %>%
            align(j=c(2,3,4,7,8,9,10,12,13,14,15), align = "center", part = "body")
          
          for(row in seq(2, nrow(speciesAssessments_n), 2)){
            
            speciesAssessments_n_ft %<>%
              merge_at(i=row, j=1:15) %>%
              bg(i=row, bg = "#EFEFEF", part = "body") %>%
              bold(i=row-1, j=1, bold=T, part="body")
          }
        }
      }else{
        speciesAssessments_n_ft <- ""
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
        elementsSpeciesOnly_g <- speciesAssessments_g_ft %>%
          delete_part(part='footer')
        
        footnotesSpeciesOnly_g <- data.frame(row=NA) %>%
          flextable() %>%
          delete_part(part='header') %>%
          border_remove() %>%
          height(i=1, height=0.1, unit="mm") %>%
          padding(padding=0, part='body')
      }
      
                            # National
      if(nrow(speciesAssessments_n) > 0){
        elementsSpeciesOnly_n <- speciesAssessments_n_ft %>%
          delete_part(part='footer')
        
        footnotesSpeciesOnly_n <- data.frame(row=NA) %>%
          flextable() %>%
          delete_part(part='header') %>%
          border_remove() %>%
          height(i=1, height=0.1, unit="mm") %>%
          padding(padding=0, part='body')
      }
    }
    
          # Ecosystem assessments
    if(triggerEcosystems > 0){
      
                # Get information
      ecosystemAssessments <- PF_ecosystems %>%
        filter(!is.na(`Criteria met`)) %>%
        mutate_at(vars(`Min site extent (km2)`, `Best site extent (km2)`, `Max site extent (km2)`, `Reference extent (km2)`), as.double) %>%
        mutate(PercentAtSite = ifelse(`Best site extent (km2)` == "-", "-", round(100 * `Best site extent (km2)`/`Reference extent (km2)`, 1))) %>%
        mutate(Blank = "") %>%
        mutate(Status = case_when(is.na(`Status in the IUCN Red List of Ecosystems`) ~ "-",
                                  language == "english" ~ paste0(`Status in the IUCN Red List of Ecosystems`, " (IUCN)"),
                                  `Status in the IUCN Red List of Ecosystems` == "Not assessed" ~ "Non évalué (UICN)",
                                  `Status in the IUCN Red List of Ecosystems` == "Data Deficient (DD)" ~ "Données insuffisantes (DD) (UICN)",
                                  `Status in the IUCN Red List of Ecosystems` == "Least Concern (LC)" ~ "Préoccupation mineure (LC) (UICN)",
                                  `Status in the IUCN Red List of Ecosystems` == "Near Threatened (NT)" ~ "Quasi menacé (NT) (UICN)",
                                  `Status in the IUCN Red List of Ecosystems` == "Vulnerable (VU)" ~ "Vulnérable (VU) (UICN)",
                                  `Status in the IUCN Red List of Ecosystems` == "Endangered (EN)" ~ "En danger (EN) (UICN)",
                                  `Status in the IUCN Red List of Ecosystems` == "Critically Endangered (CR)" ~ "En danger critique (CR) (UICN)",
                                  `Status in the IUCN Red List of Ecosystems` == "Critically Endangered (Possibly Extinct)" ~ "En danger critique (possiblement éteint) (UICN)")) %>%
        mutate(SiteExtent_Min = as.character(`Min site extent (km2)`),
               SiteExtent_Best = as.character(`Best site extent (km2)`),
               SiteExtent_Max = as.character(`Max site extent (km2)`),
               TotalExtent = as.character(`Reference extent (km2)`)) %>%
        select(`Name of ecosystem type`, `Ecosystem level`, `Ecosystem level justification`, Status, `Criteria met`, Blank, SiteExtent_Min, SiteExtent_Best, SiteExtent_Max, TotalExtent, `Data source`, PercentAtSite)
      
                # Separate global and national assessments
      ecosystemAssessments_g <- ecosystemAssessments %>%
        filter(grepl("g", `Criteria met`, fixed=T)) %>%
        mutate(`Criteria met` = gsub("g", "", `Criteria met`))
      
      ecosystemAssessments_n <- ecosystemAssessments %>%
        filter(grepl("n", `Criteria met`, fixed=T)) %>%
        mutate(`Criteria met` = gsub("n", "", `Criteria met`))
      
      if(!(nrow(ecosystemAssessments_g) + nrow(ecosystemAssessments_n)) == nrow(ecosystemAssessments)){
        convert_res[step,"Result"] <- emo::ji("prohibited")
        convert_res[step,"Message"] <- "Some ecosystem assessments are not being correctly classified as global or national assessments. This is an error with the code. Please contact Chloé and provide her with this error message."
        KBAforms[step] <- NA
        next
      }
      rm(ecosystemAssessments)
      
                # Information for the footnotes
      footnotesEcosystems_g <- ecosystemAssessments_g %>%
        select(`Ecosystem level justification`) %>%
        mutate(`Ecosystem level justification` = sapply(`Ecosystem level justification`, function(x) ifelse(substr(x, start=nchar(x), stop=nchar(x)) == ".", x, paste0(x, "."))))
      
      footnotesEcosystems_n <- ecosystemAssessments_n %>%
        select(`Ecosystem level justification`) %>%
        mutate(`Ecosystem level justification` = sapply(`Ecosystem level justification`, function(x) ifelse(substr(x, start=nchar(x), stop=nchar(x)) == ".", x, paste0(x, "."))))
      
                # Information for the main table
      ecosystemAssessments_g %<>% select(-c(`Ecosystem level justification`))
      ecosystemAssessments_n %<>% select(-c(`Ecosystem level justification`))
      
                # Assess whether min/max should be retained
                      # Remove min/max values that are identical to each other
      ecosystemAssessments_g %<>%
        mutate(SiteExtent_Min = ifelse(SiteExtent_Min == SiteExtent_Max, NA, SiteExtent_Min),
               SiteExtent_Max = ifelse(SiteExtent_Min == SiteExtent_Max, NA, SiteExtent_Max))
      
      ecosystemAssessments_n %<>%
        mutate(SiteExtent_Min = ifelse(SiteExtent_Min == SiteExtent_Max, NA, SiteExtent_Min),
               SiteExtent_Max = ifelse(SiteExtent_Min == SiteExtent_Max, NA, SiteExtent_Max))
      
                      # Check if there is only a best estimate
                            # Global
      if(sum(!is.na(ecosystemAssessments_g$SiteExtent_Min)) + sum(!is.na(ecosystemAssessments_g$SiteExtent_Max)) == 0){
        bestOnly_g <- T
        ecosystemAssessments_g %<>% select(-c(SiteExtent_Min, SiteExtent_Max))
      }else{
        bestOnly_g <- F
      }
      
                            # National
      if(sum(!is.na(ecosystemAssessments_n$SiteExtent_Min)) + sum(!is.na(ecosystemAssessments_n$SiteExtent_Max)) == 0){
        bestOnly_n <- T
        ecosystemAssessments_n %<>% select(-c(SiteExtent_Min, SiteExtent_Max))
      }else{
        bestOnly_n <- F
      }
      
                # Format flextables
                      # Ecosystem assessment - Global
      if(nrow(ecosystemAssessments_g) > 0){
        if(bestOnly_g){
          ecosystemAssessments_g_ft <- ecosystemAssessments_g %>%
            flextable()
          
          if(language == "english"){
            ecosystemAssessments_g_ft %<>%
              width(j=colnames(.), width=c(1.6,1.1,1.1,0.65,0.05,1.2,1.2,1.45,0.8)) %>%
              set_header_labels(values=list(`Name of ecosystem type` = "Ecosystem type", `Ecosystem level` = "Ecosystem level", Status = "Status", `Criteria met`="Criteria Met", Blank='', SiteExtent_Best = "Extent of ecosystem at site (km2)", TotalExtent = 'Global extent of ecosystem (km2)', `Data source` = "Source of extent data", PercentAtSite = "% of global extent at site"))
          }else{
            ecosystemAssessments_g_ft %<>%
              width(j=colnames(.), width=c(1.4,1.1,1.1,0.85,0.05,1.2,1.2,1.45,0.8)) %>%
              set_header_labels(values=list(`Name of ecosystem type` = "Type d'écosystème", `Ecosystem level` = "Niveau de l'écosystème", Status = "Statut", `Criteria met`="Critère(s) atteint(s)", Blank='', SiteExtent_Best = "Étendue de l'écosystème au site (km2)", TotalExtent = "Étendue mondiale de l'écosystème (km2)", `Data source` = "Source des données d'étendue", PercentAtSite = "% de l'étendue mondiale située au site"))
          }
          
          ecosystemAssessments_g_ft %<>%  
            align(align = "center", part="header") %>%
            font(fontname="Calibri", part="header") %>%
            fontsize(size=11, part='header') %>%
            bold(i=1, bold=T, part='header') %>%
            merge_v(part = "header") %>%
            font(fontname="Calibri", part="body") %>%
            fontsize(size=11, part='body') %>%
            hline_top(part="all") %>%
            border_remove() %>%
            hline(border = fp_border(width = 1), part="header") %>%
            hline_top(border = fp_border(width = 2), part="header") %>%
            hline_bottom(border = fp_border(width = 2), part="header") %>%
            hline_bottom(border=fp_border(width=1), part='body') %>%
            align(j=c(2,3,4,5,6,7,8,9), align = "center", part = "body")
          
        }else{
          ecosystemAssessments_g_ft <- ecosystemAssessments_g %>%
            mutate(Blank2 = "") %>%
            relocate(Blank2, .after = SiteExtent_Max) %>%
            flextable()
          
          if(language == "english"){
            ecosystemAssessments_g_ft %<>%
              width(j=colnames(.), width=c(1.6,1.2,1.3,0.65,0.05,0.7,0.7,0.7,0.05,0.7,0.7,0.8)) %>%
              set_header_labels(values=list(`Name of ecosystem type` = "Ecosystem type", `Ecosystem level` = "Ecosystem level", Status = "Status", `Criteria met`="Criteria Met", Blank='', SiteExtent_Min = "Min", SiteExtent_Best = "Best", SiteExtent_Max = "Max", Blank2 = "", TotalExtent = 'Global extent of ecosystem (km2)', `Data source` = "Source of extent data", PercentAtSite = "% of global extent at site")) %>%
              add_header_row(values = c("Ecosystem type", "Ecosystem level", "Status", "Criteria Met", "", "Extent of ecosystem at site (km2)", "", "Global extent of ecosystem (km2)", "Source of extent data", "% of global extent at site"), colwidths=c(1, 1, 1, 1, 1, 3, 1, 1, 1, 1))
          }else{
            ecosystemAssessments_g_ft %<>%
              width(j=colnames(.), width=c(1.25,1.2,1,0.85,0.05,0.7,0.75,0.7,0.05,1,0.8,0.8)) %>%
              set_header_labels(values=list(`Name of ecosystem type` = "Type d'écosystème", `Ecosystem level` = "Niveau de l'écosystème", Status = "Statut", `Criteria met`="Critère(s) atteint(s)", Blank='', SiteExtent_Min = "Min", SiteExtent_Best = "Meilleure", SiteExtent_Max = "Max", Blank2 = "", TotalExtent = "Étendue mondiale de l'écosystème (km2)", `Data source` = "Source des données d'étendue", PercentAtSite = "% de l'étendue mondiale située au site")) %>%
              add_header_row(values = c("Type d'écosystème", "Niveau de l'écosystème", "Statut", "Critère(s) atteint(s)", "", "Étendue de l'écosystème au site (km2)", "", "Étendue mondiale de l'écosystème (km2)", "Source des données d'étendue", "% de l'étendue mondiale située au site"), colwidths=c(1, 1, 1, 1, 1, 3, 1, 1, 1, 1))
          }
          
          ecosystemAssessments_g_ft %<>%  
            align(align = "center", part="header") %>%
            font(fontname="Calibri", part="header") %>%
            fontsize(size=11, part='header') %>%
            bold(i=1, bold=T, part='header') %>%
            merge_v(part = "header") %>%
            font(fontname="Calibri", part="body") %>%
            fontsize(size=11, part='body') %>%
            hline_top(part="all") %>%
            border_remove() %>%
            hline(border = fp_border(width = 1), part="header") %>%
            hline_top(border = fp_border(width = 2), part="header") %>%
            hline_bottom(border = fp_border(width = 2), part="header") %>%
            hline_bottom(border=fp_border(width=1), part='body') %>%
            align(j=c(2,3,4,5,6,7,8,9,10,11,12), align = "center", part = "body")
        }
      }else{
        ecosystemAssessments_g_ft <- ""
      }
      
                      # Ecosystem assessment - National
      if(nrow(ecosystemAssessments_n) > 0){
        if(bestOnly_n){
          ecosystemAssessments_n_ft <- ecosystemAssessments_n %>%
            flextable()
          
          if(language == "english"){
            ecosystemAssessments_n_ft %<>%
              width(j=colnames(.), width=c(1.6,1.1,1.1,0.65,0.05,1.2,1.2,1.45,0.8)) %>%
              set_header_labels(values=list(`Name of ecosystem type` = "Ecosystem type", `Ecosystem level` = "Ecosystem level", Status = "Status", `Criteria met`="Criteria Met", Blank='', SiteExtent_Best = "Extent of ecosystem at site (km2)", TotalExtent = 'National extent of ecosystem (km2)', `Data source` = "Source of extent data", PercentAtSite = "% of national extent at site"))
          }else{
            ecosystemAssessments_n_ft %<>%
              width(j=colnames(.), width=c(1.4,1.1,1.1,0.85,0.05,1.2,1.2,1.45,0.8)) %>%
              set_header_labels(values=list(`Name of ecosystem type` = "Type d'écosystème", `Ecosystem level` = "Niveau de l'écosystème", Status = "Statut", `Criteria met`="Critère(s) atteint(s)", Blank='', SiteExtent_Best = "Étendue de l'écosystème au site (km2)", TotalExtent = "Étendue nationale de l'écosystème (km2)", `Data source` = "Source des données d'étendue", PercentAtSite = "% de l'étendue nationale située au site"))
          }
          
          ecosystemAssessments_n_ft %<>%  
            align(align = "center", part="header") %>%
            font(fontname="Calibri", part="header") %>%
            fontsize(size=11, part='header') %>%
            bold(i=1, bold=T, part='header') %>%
            merge_v(part = "header") %>%
            font(fontname="Calibri", part="body") %>%
            fontsize(size=11, part='body') %>%
            hline_top(part="all") %>%
            border_remove() %>%
            hline(border = fp_border(width = 1), part="header") %>%
            hline_top(border = fp_border(width = 2), part="header") %>%
            hline_bottom(border = fp_border(width = 2), part="header") %>%
            hline_bottom(border=fp_border(width=1), part='body') %>%
            align(j=c(2,3,4,5,6,7,8,9), align = "center", part = "body")
          
        }else{
          ecosystemAssessments_n_ft <- ecosystemAssessments_n %>%
            mutate(Blank2 = "") %>%
            relocate(Blank2, .after = SiteExtent_Max) %>%
            flextable()
          
          if(language == "english"){
            ecosystemAssessments_n_ft %<>%
              width(j=colnames(.), width=c(1.6,1.2,1.3,0.65,0.05,0.7,0.7,0.7,0.05,0.7,0.7,0.8)) %>%
              set_header_labels(values=list(`Name of ecosystem type` = "Ecosystem type", `Ecosystem level` = "Ecosystem level", Status = "Status", `Criteria met`="Criteria Met", Blank='', SiteExtent_Min = "Min", SiteExtent_Best = "Best", SiteExtent_Max = "Max", Blank2 = "", TotalExtent = 'National extent of ecosystem (km2)', `Data source` = "Source of extent data", PercentAtSite = "% of national extent at site")) %>%
              add_header_row(values = c("Ecosystem type", "Ecosystem level", "Status", "Criteria Met", "", "Extent of ecosystem at site (km2)", "", "National extent of ecosystem (km2)", "Source of extent data", "% of national extent at site"), colwidths=c(1, 1, 1, 1, 1, 3, 1, 1, 1, 1))
          }else{
            ecosystemAssessments_n_ft %<>%
              width(j=colnames(.), width=c(1.25,1.2,1,0.85,0.05,0.7,0.75,0.7,0.05,1,0.8,0.8)) %>%
              set_header_labels(values=list(`Name of ecosystem type` = "Type d'écosystème", `Ecosystem level` = "Niveau de l'écosystème", Status = "Statut", `Criteria met`="Critère(s) atteint(s)", Blank='', SiteExtent_Min = "Min", SiteExtent_Best = "Meilleure", SiteExtent_Max = "Max", Blank2 = "", TotalExtent = "Étendue nationale de l'écosystème (km2)", `Data source` = "Source des données d'étendue", PercentAtSite = "% de l'étendue nationale située au site")) %>%
              add_header_row(values = c("Type d'écosystème", "Niveau de l'écosystème", "Statut", "Critère(s) atteint(s)", "", "Étendue de l'écosystème au site (km2)", "", "Étendue nationale de l'écosystème (km2)", "Source des données d'étendue", "% de l'étendue nationale située au site"), colwidths=c(1, 1, 1, 1, 1, 3, 1, 1, 1, 1))
          }
          
          ecosystemAssessments_n_ft %<>%  
            align(align = "center", part="header") %>%
            font(fontname="Calibri", part="header") %>%
            fontsize(size=11, part='header') %>%
            bold(i=1, bold=T, part='header') %>%
            merge_v(part = "header") %>%
            font(fontname="Calibri", part="body") %>%
            fontsize(size=11, part='body') %>%
            hline_top(part="all") %>%
            border_remove() %>%
            hline(border = fp_border(width = 1), part="header") %>%
            hline_top(border = fp_border(width = 2), part="header") %>%
            hline_bottom(border = fp_border(width = 2), part="header") %>%
            hline_bottom(border=fp_border(width=1), part='body') %>%
            align(j=c(2,3,4,5,6,7,8,9,10,11,12), align = "center", part = "body")
        }
      }else{
        ecosystemAssessments_n_ft <- ""
      }
      
                       # Add footnotes, with formatted hyperlinks
                            # Global
      if(nrow(ecosystemAssessments_g) > 0){
        footnote <- 0
        
        for(i in 1:nrow(ecosystemAssessments_g)){
          col <- which(grepl("http", footnotesEcosystems_g[i,]), arr.ind = TRUE)
          
          for(c in 1:ncol(footnotesEcosystems_g)){
            string <- footnotesEcosystems_g[i,c]
            
            if(!is.na(string)){
              footnote <- footnote+1
              
              # If there's a link in the footnote
              if(c %in% col){
                urls <- str_locate_all(string, "http")[[1]][,1]
                urlIDs <- paste0("url", urls)
                spaces <- str_locate_all(string, " ")[[1]][,1]
                if(length(spaces) == 0){
                  spaces <- -1
                }
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
                  
                call_final <- paste0("ecosystemAssessments_g_ft %<>% footnote(i=i, j=2, value=as_paragraph(", call_final,"), ref_symbols=as.integer(footnote))")
                
                # Evaluate call
                eval(parse(text=call_final))
                
                # If there is no link in the footnote
              }else{
                ecosystemAssessments_g_ft %<>% footnote(i=i, j=2, value=as_paragraph(as.character(string)), ref_symbols=as.integer(footnote))
              }
            }
          }
        }
      }
      
                            # National
      if(nrow(ecosystemAssessments_n) > 0){
        footnote <- 0
        
        for(i in 1:nrow(ecosystemAssessments_n)){
          col <- which(grepl("http", footnotesEcosystems_n[i,]), arr.ind = TRUE)
          
          for(c in 1:ncol(footnotesEcosystems_n)){
            string <- footnotesEcosystems_n[i,c]
            
            if(!is.na(string)){
              footnote <- footnote+1
              
              # If there's a link in the footnote
              if(c %in% col){
                urls <- str_locate_all(string, "http")[[1]][,1]
                urlIDs <- paste0("url", urls)
                spaces <- str_locate_all(string, " ")[[1]][,1]
                if(length(spaces) == 0){
                  spaces <- -1
                }
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
                
                call_final <- paste0("ecosystemAssessments_n_ft %<>% footnote(i=i, j=2, value=as_paragraph(", call_final,"), ref_symbols=as.integer(footnote))")
                
                # Evaluate call
                eval(parse(text=call_final))
                
                # If there is no link in the footnote
              }else{
                
                ecosystemAssessments_n_ft %<>% footnote(i=i, j=2, value=as_paragraph(as.character(string)), ref_symbols=as.integer(footnote))
              }
            }
          }
        }
      }
      
                      # Add padding
                            # Global
      if(nrow(ecosystemAssessments_g) > 0){
        ecosystemAssessments_g_ft %<>%
          padding(padding.top = 10, part='footer') %>%
          font(fontname='Calibri', part='footer')
      }
      
                            # National
      if(nrow(ecosystemAssessments_n) > 0){
        ecosystemAssessments_n_ft %<>%
          padding(padding.top = 10, part='footer') %>%
          font(fontname='Calibri', part='footer')
      }
      
                      # Prepare final tables
                            # Global
      if(nrow(ecosystemAssessments_g) > 0){
        elementsEcosystemsOnly_g <- ecosystemAssessments_g_ft %>%
          delete_part(part='footer')
        
        footnotesEcosystemsOnly_g <- ecosystemAssessments_g_ft %>%
          delete_part(part='header') %>%
          delete_part(part='body') %>%
          bg(bg = "#EFEFEF", part = "footer")
      }
      
                            # National
      if(nrow(ecosystemAssessments_n) > 0){
        elementsEcosystemsOnly_n <- ecosystemAssessments_n_ft %>%
          delete_part(part='footer')
        
        footnotesEcosystemsOnly_n <- ecosystemAssessments_n_ft %>%
          delete_part(part='header') %>%
          delete_part(part='body') %>%
          bg(bg = "#EFEFEF", part = "footer")
      }
    }
    
          # Trigger Elements summary - Species
    if(triggerSpecies > 0){
      
      elementsSummary_spp <- PF_species %>%
        filter(!is.na(`Criteria met`)) %>%
        select(`Common name`, `Scientific name`) %>%
        unique() %>%
        pivot_longer(., cols=c(`Common name`, `Scientific name`), names_to="Type") %>%
        select(-Type) %>%
        t() %>%
        data.frame() %>%
        mutate(Prefix = ifelse(language=="english",
                               paste0(as.character("\u25CF"), " Species: "),
                               paste0(as.character("\u25CF"), " Espèce(s) : ")))
      elementsSummary_spp <- elementsSummary_spp[,c(ncol(elementsSummary_spp), 1:(ncol(elementsSummary_spp)-1))]
      
      elementsSummary_spp_ft <- flextable(elementsSummary_spp, col_keys = c("Blank", "Label"), defaults=list(fontname="Calibri", font.size=11)) %>%
        width(j=colnames(.), width=c(0.3, 9))
      
      extraCall <- ""
      if(ncol(elementsSummary_spp) > 3){
        
        # Keep only columns with common names
        spp <- 4:ncol(elementsSummary_spp)
        spp <- spp[lapply(spp, "%%", 2) == 0]
        
        for(i in spp){
          
          if(!i == spp[length(spp)]){
            
            if(elementsSummary_spp[i] == elementsSummary_spp[i+1]){
              extraCall <- paste0(extraCall, ", as_chunk(x=', '), as_chunk(x=X", i-1, ")")
            }else{
              extraCall <- paste0(extraCall, ", as_chunk(x=', '), as_chunk(x=X", i-1, "), as_chunk(x=' ('), as_chunk(x=X", i, ", props=fp_text(font.size=11, font.family='Calibri', italic = T)), as_chunk(x=')')")
            }
            
          }else{
            
            if(elementsSummary_spp[i] == elementsSummary_spp[i+1]){
              extraCall <- ifelse(language == "english",
                                  paste0(extraCall, ", as_chunk(x=' and '), as_chunk(x=X", i-1, ")"),
                                  paste0(extraCall, ", as_chunk(x=' et '), as_chunk(x=X", i-1, ")"))
            }else{
              extraCall <- ifelse(language == "english",
                                  paste0(extraCall, ", as_chunk(x=' and '), as_chunk(x=X", i-1, "), as_chunk(x=' ('), as_chunk(x=X", i, ", props=fp_text(font.size=11, font.family='Calibri', italic = T)), as_chunk(x=')')"),
                                  paste0(extraCall, ", as_chunk(x=' et '), as_chunk(x=X", i-1, "), as_chunk(x=' ('), as_chunk(x=X", i, ", props=fp_text(font.size=11, font.family='Calibri', italic = T)), as_chunk(x=')')"))
            }
          }
        }
      }
      
      if(elementsSummary_spp[2] == elementsSummary_spp[3]){
        compose_call <- paste0("elementsSummary_spp_ft %<>% compose(j='Label', value=as_paragraph(as_chunk(Prefix), as_chunk(x=X1)", 
                               extraCall,
                               "))")
      }else{
        compose_call <- paste0("elementsSummary_spp_ft %<>% compose(j='Label', value=as_paragraph(as_chunk(Prefix), as_chunk(x=X1), as_chunk(x=' ('), as_chunk(x=X2, props=fp_text(font.size=11, font.family='Calibri', italic = T)), as_chunk(x=')')", 
                               extraCall,
                               "))")
      }
      eval(parse(text=compose_call))
      
      elementsSummary_spp_ft %<>%
        delete_part(part='header') %>%
        border_remove() %>%
        align(j=2, align = "left", part = "body") %>%
        font(fontname = "Calibri", part="body")
    }
    
            # Trigger Elements summary - Ecosystems
    if(triggerEcosystems > 0){
      elementsSummary_eco <- PF_ecosystems %>%
        filter(!is.na(`Criteria met`)) %>%
        select(`Name of ecosystem type`) %>%
        unique() %>%
        t() %>%
        data.frame() %>%
        mutate(Prefix = ifelse(language=="english",
                               paste0(as.character("\u25CF"), " Ecosystem type(s): "),
                               paste0(as.character("\u25CF"), " Type(s) d'écosystème(s) : ")))
      elementsSummary_eco <- elementsSummary_eco[,c(ncol(elementsSummary_eco), 1:(ncol(elementsSummary_eco)-1))]
      
      if(ncol(elementsSummary_eco) == 3){
        
        elementsSummary_eco %<>%
          mutate(X1 = paste(.[,2:3], collapse=ifelse(language=="english", " and ", " et "))) 
        
      }else if(ncol(elementsSummary_eco) > 3){
        
        elementsSummary_eco %<>%
          mutate(X1 = paste(paste(.[,2:(ncol(.)-1)], collapse=", "), .[,ncol(.)], collapse=ifelse(language=="english", " and ", " et ")))
      }
      
      elementsSummary_eco %<>%
        mutate(Prefix = paste0(Prefix, X1))
      
      elementsSummary_eco_ft <- flextable(elementsSummary_eco, col_keys = c("Blank", "Label"), defaults=list(fontname="Calibri", font.size=11)) %>%
        width(j=colnames(.), width=c(0.3, 9)) %>%
        compose(j='Label', value=as_paragraph(as_chunk(Prefix))) %>%
        delete_part(part='header') %>%
        border_remove() %>%
        align(j=2, align = "left", part = "body") %>%
        font(fontname = "Calibri", part="body")
    }
    
          # Subtitle (cover page) - Species
    if(triggerSpecies > 0){
      
      elementsSummary_spp %<>% select(-Prefix)
      subtitle_spp_ft <- flextable(elementsSummary_spp, col_keys = "Label", defaults=list(fontname="Calibri", font.size=12, color='#5A5A5A')) %>%
        width(j=colnames(.), width=c(9))
      
      extraCall <- ""
      if(ncol(elementsSummary_spp) > 2){
        
        # Keep only columns with scientific names
        spp <- 4:ncol(elementsSummary_spp)
        spp <- spp[lapply(spp, "%%", 2) == 0]
        
        for(i in spp){
          
          if(!i == spp[length(spp)]){
            
            if(elementsSummary_spp[i-1] == elementsSummary_spp[i]){
              extraCall <- paste0(extraCall, ", as_chunk(x=', ', props=fp_text(font.size=12, font.family='Calibri', color='#5A5A5A')), as_chunk(x=X", i-1, ", props=fp_text(font.size=12, font.family='Calibri', color='#5A5A5A'))")
            }else{
              extraCall <- paste0(extraCall, ", as_chunk(x=', ', props=fp_text(font.size=12, font.family='Calibri', color='#5A5A5A')), as_chunk(x=X", i-1, ", props=fp_text(font.size=12, font.family='Calibri', color='#5A5A5A')), as_chunk(x=' (', props=fp_text(font.size=12, font.family='Calibri', color='#5A5A5A')), as_chunk(x=X", i, ", props=fp_text(font.size=12, font.family='Calibri', italic = T, color='#5A5A5A')), as_chunk(x=')', props=fp_text(font.size=12, font.family='Calibri', color='#5A5A5A'))")
            }
            
          }else{
            
            if(elementsSummary_spp[i-1] == elementsSummary_spp[i]){
              extraCall <- ifelse(language=="english",
                                  paste0(extraCall, ", as_chunk(x=' and ', props=fp_text(font.size=12, font.family='Calibri', color='#5A5A5A')), as_chunk(x=X", i-1, ", props=fp_text(font.size=12, font.family='Calibri', color='#5A5A5A'))"),
                                  paste0(extraCall, ", as_chunk(x=' et ', props=fp_text(font.size=12, font.family='Calibri', color='#5A5A5A')), as_chunk(x=X", i-1, ", props=fp_text(font.size=12, font.family='Calibri', color='#5A5A5A'))"))
            }else{
              extraCall <- ifelse(language=="english",
                                  paste0(extraCall, ", as_chunk(x=' and ', props=fp_text(font.size=12, font.family='Calibri', color='#5A5A5A')), as_chunk(x=X", i-1, ", props=fp_text(font.size=12, font.family='Calibri', color='#5A5A5A')), as_chunk(x=' (', props=fp_text(font.size=12, font.family='Calibri', color='#5A5A5A')), as_chunk(x=X", i, ", props=fp_text(font.size=12, font.family='Calibri', italic = T, color='#5A5A5A')), as_chunk(x=')', props=fp_text(font.size=12, font.family='Calibri', color='#5A5A5A'))"),
                                  paste0(extraCall, ", as_chunk(x=' et ', props=fp_text(font.size=12, font.family='Calibri', color='#5A5A5A')), as_chunk(x=X", i-1, ", props=fp_text(font.size=12, font.family='Calibri', color='#5A5A5A')), as_chunk(x=' (', props=fp_text(font.size=12, font.family='Calibri', color='#5A5A5A')), as_chunk(x=X", i, ", props=fp_text(font.size=12, font.family='Calibri', italic = T, color='#5A5A5A')), as_chunk(x=')', props=fp_text(font.size=12, font.family='Calibri', color='#5A5A5A'))"))
            }
          }
        }
      }
      
      if(elementsSummary_spp[1] == elementsSummary_spp[2]){
        
        if(language == "english"){
          compose_call <- paste0("subtitle_spp_ft %<>% compose(j='Label', value=as_paragraph(as_chunk(x=paste('Species:', X1), props=fp_text(font.size=12, font.family='Calibri', color='#5A5A5A'))", 
                                 extraCall,
                                 "))")
        }else{
          compose_call <- paste0("subtitle_spp_ft %<>% compose(j='Label', value=as_paragraph(as_chunk(x=paste('Espèce(s) :', X1), props=fp_text(font.size=12, font.family='Calibri', color='#5A5A5A'))", 
                                 extraCall,
                                 "))")
        }
        
      }else{
        
        if(language == "english"){
          compose_call <- paste0("subtitle_spp_ft %<>% compose(j='Label', value=as_paragraph(as_chunk(x=paste('Species:', X1), props=fp_text(font.size=12, font.family='Calibri', color='#5A5A5A')), as_chunk(x=' (', props=fp_text(font.size=12, font.family='Calibri', color='#5A5A5A')), as_chunk(x=X2, props=fp_text(font.size=12, font.family='Calibri', italic = T, color='#5A5A5A')), as_chunk(x=')', props=fp_text(font.size=12, font.family='Calibri', color='#5A5A5A'))", 
                                 extraCall,
                                 "))")
        }else{
          compose_call <- paste0("subtitle_spp_ft %<>% compose(j='Label', value=as_paragraph(as_chunk(x=paste('Espèce(s) :', X1), props=fp_text(font.size=12, font.family='Calibri', color='#5A5A5A')), as_chunk(x=' (', props=fp_text(font.size=12, font.family='Calibri', color='#5A5A5A')), as_chunk(x=X2, props=fp_text(font.size=12, font.family='Calibri', italic = T, color='#5A5A5A')), as_chunk(x=')', props=fp_text(font.size=12, font.family='Calibri', color='#5A5A5A'))", 
                                 extraCall,
                                 "))")
        }
      }
      eval(parse(text=compose_call))
      
      subtitle_spp_ft %<>%
        delete_part(part='header') %>%
        border_remove() %>%
        align(j=1, align = "left", part = "body") %>%
        fontsize(size=12, part='body')
    }
    
          # Subtitle (cover page) - Ecosystems
    if(triggerEcosystems > 0){
      
      elementsSummary_eco %<>%
        select(Prefix) %>%
        mutate(Prefix = gsub(as.character("\u25CF "), "", Prefix))
      
      subtitle_eco_ft <- flextable(elementsSummary_eco, col_keys = "Label", defaults=list(fontname="Calibri", font.size=12, color='#5A5A5A')) %>%
        width(j=colnames(.), width=c(9)) %>%
        compose(j='Label', value=as_paragraph(as_chunk(x=Prefix, props=fp_text(font.size=12, font.family='Calibri', color='#5A5A5A')))) %>%
        delete_part(part='header') %>%
        border_remove() %>%
        align(j=1, align = "left", part = "body") %>%
        fontsize(size=12, part='body')
    }
    
          # Technical Review
    if(reviewStage == "general"){
      technicalReview_ft <- PF_technicalReview %>%
        select(-`Description of role`) %>%
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
      
    }else{
      technicalReview_ft <- PF_technicalReview %>%
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
    }
    
          # General Review
    if(ncol(PF_generalReview) == 3){
      widths <- c(2.4,3.6,3)
    }else{
      widths <- c(1.4,2,2,3.6)
    }
   
    generalReview_ft <- PF_generalReview %>%
      flextable() %>%
      width(j=colnames(.), width=widths) %>%
      align(align = "center", part="header") %>%
      font(fontname="Calibri", part="header") %>%
      fontsize(size=11, part='header') %>%
      bold(i=1, bold=T, part='header') %>%
      merge_v(part = "header") %>%
      font(fontname="Calibri", part="body") %>%
      fontsize(size=11, part='body') %>%
      hline_top(part="all")
    
          # Reviewers that did not provide feedback
    noFeedback_ft <- PF_noReview %>%
      {ifelse((. == "None") & (language=="french"), "Aucun", .)} %>%
      as.data.frame() %>%
      flextable() %>%
      width(j=colnames(.), width=9) %>%
      delete_part(., part = "header") %>%
      font(fontname="Calibri", part="body") %>%
      fontsize(size=11, part='body') %>%
      border_remove()
    
          # Additional Site Information
    additionalInfo <- data.frame(Type = character(),
                                 Value = character(),
                                 stringsAsFactors = F)
    
                # Nomination rationale
    if(language == "english"){
      additionalInfo[1, ] <- c("Rationale for site nomination", nominationRationale)
    }else{
      additionalInfo[1, ] <- c("Justification de la proposition", nominationRationale)
    }
    
                # Site history
    if(!is.na(siteHistory)){
      
      if(language == "english"){
        additionalInfo[2, ] <- c("Site history", siteHistory)
      }else{
        additionalInfo[2, ] <- c("Historique du site", siteHistory)
      }
    }
    
                # Assessed elements that did not meet KBA criteria
    if(!reviewStage == "general"){
      speciesNotTriggers <- PF_species %>%
        filter(is.na(`Criteria met`) & !`Scientific name` %in% PF_species$`Scientific name`[which(!is.na(PF_species$`Criteria met`))]) %>%
        select(`Common name`, `Scientific name`) %>%
        distinct() %>%
        mutate(Label = ifelse(!`Common name` == `Scientific name`, paste0(`Common name`, " (<i>", `Scientific name`, "</i>)"), `Common name`)) %>%
        pull(Label) %>%
        paste(., collapse=", ")
      
      ecosystemsNotTriggers <- PF_ecosystems %>%
        filter(is.na(`Criteria met`) & !`Name of ecosystem type` %in% PF_ecosystems$`Name of ecosystem type`[which(!is.na(PF_ecosystems$`Criteria met`))]) %>%
        pull(`Name of ecosystem type`) %>%
        unique() %>%
        paste(., collapse=", ")
      
      if(language == "english"){
        additionalInfo[nrow(additionalInfo)+1, ] <- c("Biodiversity elements that were assessed but did not meet KBA criteria", ifelse((speciesNotTriggers == "") & (ecosystemsNotTriggers == ""), "-", paste(ifelse(!ecosystemsNotTriggers == "", paste("Ecosystem type(s):", ecosystemsNotTriggers), ""), ifelse(!speciesNotTriggers == "", paste("Species:", speciesNotTriggers), ""), sep=ifelse((!speciesNotTriggers == "") & (!ecosystemsNotTriggers == ""), "\n\n", ""))))
      }else{
        additionalInfo[nrow(additionalInfo)+1, ] <- c("Éléments de biodiversité évalués qui n’atteignent pas les critères KBA", ifelse((speciesNotTriggers == "") & (ecosystemsNotTriggers == ""), "-", paste(ifelse(!ecosystemsNotTriggers == "", paste("Type(s) d'écosystème(s) :", ecosystemsNotTriggers), ""), ifelse(!speciesNotTriggers == "", paste("Espèces :", speciesNotTriggers), ""), sep=ifelse((!speciesNotTriggers == "") & (!ecosystemsNotTriggers == ""), "\n\n", ""))))
      }
    }
    
                # Additional biodiversity
    if(language == "english"){
      additionalInfo[nrow(additionalInfo)+1, ] <- c("Additional biodiversity at the site", ifelse(is.na(additionalBiodiversity), "-", additionalBiodiversity))
    }else{
      additionalInfo[nrow(additionalInfo)+1, ] <- c("Biodiversité additionnelle au site", ifelse(is.na(additionalBiodiversity), "-", additionalBiodiversity))
    }
    
                # Customary jurisdiction at site
    if(language == "english"){
      additionalInfo[nrow(additionalInfo)+1, ] <- c("Customary jurisdiction at site", ifelse(is.na(customaryJurisdiction), "-", customaryJurisdiction))
    }else{
      additionalInfo[nrow(additionalInfo)+1, ] <- c("Juridiction coutumière au site", ifelse(is.na(customaryJurisdiction), "-", customaryJurisdiction))
    }
    
                # Customary jurisdiction source
    if(!PF_formVersion %in% c(1, 1.1)){
      
      if(language == "english"){
        additionalInfo[nrow(additionalInfo)+1, ] <- c("Source of customary jurisdiction information", ifelse(is.na(customaryJurisdictionSource), "-", customaryJurisdictionSource))
      }else{
        additionalInfo[nrow(additionalInfo)+1, ] <- c("Source de l'information sur la juridiction coutumière", ifelse(is.na(customaryJurisdictionSource), "-", customaryJurisdictionSource))
      }
    }
    
                # Conservation
    additionalInfo[nrow(additionalInfo)+1, ] <- c("Conservation", ifelse(is.na(conservation), "-", conservation))
    
                # Ongoing conservation actions
    ongoingActions <- PF_actions %>%
      filter(Ongoing == "TRUE") %>%
      pull(Action) %>%
      lapply(., function(x) ifelse(x=="None", x, substr(x, start=5, stop=nchar(x)))) %>%
      unlist()
    
    if(language == "french"){
      ongoingActions %<>%
        lapply(., function(x) xwalk_conservationAction[which(xwalk_conservationAction$ConservationAction_EN == x), "ConservationAction_FR"]) %>%
        unlist()
    }
    
    ongoingActions %<>%
      sort() %>%
      paste(., collapse="; ")
    
    if(language == "english"){
      additionalInfo[nrow(additionalInfo)+1, ] <- c("Ongoing conservation actions", ifelse((length(ongoingActions) == 0) | (ongoingActions == ""), "None", ongoingActions))
    }else{
      additionalInfo[nrow(additionalInfo)+1, ] <- c("Actions de conservation en cours", ifelse((length(ongoingActions) == 0) | (ongoingActions == ""), "Aucune", ongoingActions))
    }
    
                # Ongoing threats
    if(!PF_noThreats){
      threatText <- PF_threats %>%
        pull(`Level 1`) %>%
        unique() %>%
        substr(., start=3, stop=nchar(.)) %>%
        trimws()
      
      if(language == "french"){
        threatText %<>%
          lapply(., function(x) xwalk_threat[which(xwalk_threat$Threat_EN == x), "Threat_FR"]) %>%
          unlist %>%
          trimws()
      }
      
      threatText %<>%
        sort() %>%
        paste(., collapse='; ')
      
    }else{
      threatText <- "-"
    }
    
    if(language == "english"){
      additionalInfo[nrow(additionalInfo)+1, ] <- c("Ongoing threats", threatText)
    }else{
      additionalInfo[nrow(additionalInfo)+1, ] <- c("Menaces actuelles", threatText)
    }
    
                # Conservation actions needed
    neededActions <- PF_actions %>%
      filter(Needed == "TRUE") %>%
      pull(Action) %>%
      lapply(., function(x) ifelse(x=="None", x, substr(x, start=5, stop=nchar(x)))) %>%
      unlist()
    
    if(language == "french"){
      neededActions %<>% lapply(., function(x) xwalk_conservationAction[which(xwalk_conservationAction$ConservationAction_EN == x), "ConservationAction_FR"]) %>%
        unlist()
    }
    
    neededActions %<>%
      sort() %>%
      paste(., collapse="; ")
    
    if(language == "english"){
      additionalInfo[nrow(additionalInfo)+1, ] <- c("Conservation actions needed", ifelse((length(neededActions) == 0) | (neededActions == ""), "-", neededActions))
    }else{
      additionalInfo[nrow(additionalInfo)+1, ] <- c("Actions de conservation nécessaires", ifelse((length(neededActions) == 0) | (neededActions == ""), "-", neededActions))
    }
    
                # Make it a flextable
    additionalInfo_ft <- additionalInfo[0,] %>%
      flextable()
    
    for(row in 1:nrow(additionalInfo)){
      
      additionalInfo_ft %<>% add_body_row(., top=F, value=additionalInfo[row,])
      
      if(grepl("<i>", additionalInfo[row,2], fixed = T)){
        
        iStart <- gregexpr("<i>", additionalInfo[row,2], fixed=T)[[1]]
        iEnd <- gregexpr("</i>", additionalInfo[row,2], fixed=T)[[1]]
        
        compose_call <- ""
        
        for(it in 1:length(iStart)){
          
          normalStart <- ifelse(it==1, 1, italicEnd+5)
          normalStop <- iStart[it]-1
          
          italicStart <- iStart[it]+3
          italicEnd <- iEnd[it]-1
          
          compose_call <- paste0(compose_call, ifelse(compose_call=="", "", ", "), "as_chunk(x=substr(additionalInfo[row,2], ", normalStart, ", ", normalStop, "))", ", as_chunk(x=substr(additionalInfo[row,2], ", italicStart, ", ", italicEnd, "), props=fp_text(font.size=11, font.family='Calibri', italic=T))")
          
        }
        
        if(!(italicEnd+4) == nchar(additionalInfo[row,2])){
          
          normalStart <- italicEnd+5
          normalStop <- nchar(additionalInfo[row,2])
          
          compose_call <- paste0(compose_call, ", as_chunk(x=substr(additionalInfo[row,2], ", normalStart, ", ", normalStop, "))")
        }
        
        call_final <- paste0("additionalInfo_ft %<>% compose(i=row, j='Value', value=as_paragraph(", compose_call, "))")
        eval(parse(text=call_final))
      }
    }
    
    additionalInfo_ft %<>%
      width(j=colnames(.), width=c(3,6)) %>%
      font(fontname="Calibri", part="body") %>%
      fontsize(size=11, part='body') %>%
      delete_part(part="header") %>%
      theme_zebra() %>%
      align(j=c(1,2), align='left', part='body') %>%
      bold(j=1)
    
          # Citations
    if(nrow(PF_citations) == 0){
      PF_citations[1,] <- c("", "No references provided.", "", "", "")
    }
    
    citations_ft <- PF_citations %>%
      arrange(`Long citation`) %>%
      select(`Long citation`) %>%
      flextable() %>%
      delete_part(part = "header") %>%
      width(j=colnames(.), width=9) %>%
      border_remove() %>%
      font(fontname="Calibri", part="body")
    
    #### Save KBA summary
    # Get Word template
          # Get template list
    templateList <- "https://drive.google.com/drive/folders/1gmQJpzWJUK-udCCCovpKRxFNYbxz06UR" %>%
      drive_get(as_id(.)) %>%
      drive_ls(.)
    
          # Get template name
    templateName <- paste0("Template_", reviewStage, "_", paste0(c("gE", "nE", "gS", "nS")[which(c(gE, nE, gS, nS))], collapse=""), "_", language, ".docx")
    
          # Download template
    drive_download(paste0("https://docs.google.com/document/d/", templateList$id[which(templateList$name == templateName)]), overwrite = TRUE)
    
    # List all flextables
          # All cases
    FT <- list(criteria = criteriaInfo_ft,
               technicalReview = technicalReview_ft,
               generalReview = generalReview_ft,
               noFeedback = noFeedback_ft,
               additionalInfo = additionalInfo_ft,
               references = citations_ft,
               siteDescription = siteDescription_ft,
               delineationRationale = delineationRationale_ft)
    
          # Dependent on types of triggers
                # Ecosystems
    if(gE | nE){
      FT <- append(FT, list(subtitle_eco = subtitle_eco_ft,
                            triggerElements_eco = elementsSummary_eco_ft))
    }
    
    if(gE){
      FT <- append(FT, list(elements_g_eco = elementsEcosystemsOnly_g,
                            elementFootnotes_g_eco = footnotesEcosystemsOnly_g))
    }
    
    if(nE){
      FT <- append(FT, list(elements_n_eco = elementsEcosystemsOnly_n,
                            elementFootnotes_n_eco = footnotesEcosystemsOnly_n))
    }
    
                # Species
    if(gS | nS){
      FT <- append(FT, list(subtitle_spp = subtitle_spp_ft,
                            triggerElements_spp = elementsSummary_spp_ft))
    }
    
    if(gS){
      FT <- append(FT, list(elements_g_spp = elementsSpeciesOnly_g,
                            elementFootnotes_g_spp = footnotesSpeciesOnly_g))
    }
    
    if(nS){
      FT <- append(FT, list(elements_n_spp = elementsSpeciesOnly_n,
                            elementFootnotes_n_spp = footnotesSpeciesOnly_n))
    }
   
    # Compute document name
    reviewStageLabel <- ifelse(reviewStage == "technical",
                               "TR",
                               ifelse(reviewStage == "general",
                                      "GR",
                                      "SC"))
    
    if(language == "english"){
      doc <- paste0("KBASummary_", reviewStageLabel, "_", str_replace_all(string=nationalName, pattern=c(":| |\\(|\\)|/"), repl=""), "_", Sys.Date(), ".docx")
      
    }else{
      doc <- paste0("SommaireKBA_", reviewStageLabel, "_", str_replace_all(string=nationalName, pattern=c(":| |\\(|\\)|/"), repl=""), "_", Sys.Date(), ".docx")
    }
    
    # Save
    doc <- renderInlineCode(templateName, doc)
    Sys.sleep(5)
    doc <- body_add_flextables(doc, doc, FT)
    
    KBAforms[step] <- doc
    
    # If the previous doc object was created, than the conversion was a success
    success <- TRUE
    
    # Compute the successful conversion into the table
    if(success){
      convert_res[step,"Result"] <- emo::ji("check")
      
      if(triggerSpecies > 0){
        for(i in 1:nrow(PF_species)){
          
          if(((!is.na(PF_species$display_taxonname[i]) && (PF_species$display_taxonname[i] == "No")) | (!is.na(PF_species$display_taxonomicgroup[i]) && (PF_species$display_taxonomicgroup[i] == "No")) | (!is.na(PF_species$display_assessmentinfo[i]) && (PF_species$display_assessmentinfo[i] == "No"))) & (!reviewStage == "general")){
            convert_res[step, "Message"] <- "WARNING: Contains unredacted sensitive information"
          }
        }
      }
    }
  }
  
  convert_res <<- convert_res
  KBAforms_doc <<- KBAforms
}

form_conversion <- function(KBAforms, reviewStage, language, app){
  
  if(app){
    withProgress(message = "Converting forms", value = 0, summary(KBAforms, reviewStage, language, app))
  }else{
    summary(KBAforms, reviewStage, language, app)
  }
  
  # List to store the summaries AND the result table that will be displayed on the Shiny app 
  list_item <- list() # list to stock the summaries and a dataframe to see if it's a success or not
  list_item[[1]] <- KBAforms_doc
  list_item[[2]] <- convert_res
  
  return(list_item)
}
