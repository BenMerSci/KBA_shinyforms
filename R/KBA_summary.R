form_conversion <- function(KBAforms, reviewStage, language){
  
  withProgress(message = "Converting forms", value = 0, {
    
    # Options
    options(scipen = 999)
    
    # Load crosswalks
          # Assessment Paramter
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
    
    # Load species list
    if(language == "french"){
      googledrive::drive_download("https://docs.google.com/spreadsheets/d/1R2ILLvyGMqRL8S9pfZdYIeBKXlyzckKQ", overwrite = T)
      masterSpeciesList <- read_excel("Ref_Species.xlsx", sheet=2)
      write_excel_csv(masterSpeciesList, file="Ref_Species.csv")
      masterSpeciesList <- read_csv("Ref_Species.csv")
    }
    
    # Create a dataframe to store the success/failure state of each conversion
    convert_res <- data.frame(matrix(ncol=3))
    colnames(convert_res) <- c("Name","Result","Message")
    
    #### Prepare the Summary(ies) ####
    for(step in 1:length(KBAforms)){
    
      if(!grepl(".xlsm", KBAforms[step], fixed =  TRUE)){
        convert_res[step,"Result"] <- emo::ji("prohibited")
        convert_res[step, "Message"] <- paste(KBAforms[step], "is not a KBA proposal form")
        KBAforms[step] <- NA
        next
      }
    
      incProgress(1/length(KBAforms), detail = paste("form number ", step))
    
      success <- FALSE # Set success to FALSE
      
      # Load KBA Canada Proposal Form
            # Load full workbook
      wb <- loadWorkbook(KBAforms[step])
      
      if(sum(c("HOME", "1. PROPOSER", "2. SITE", "3. SPECIES","4. ECOSYSTEMS & C", "5. THREATS", "6. REVIEW", "7. CITATIONS", "8. CHECK") %in% getSheetNames(KBAforms[step])) != 9) {
        convert_res[step, "Result"] <- emo::ji("prohibited")
        convert_res[step, "Message"] <- paste(KBAforms[step], "is not a KBA Canada proposal form. If you need a summary for a Legacy single-site form, contact Chloé.")
        KBAforms[step] <- NA
        next
      }
            # Visible sheets
      home <- read.xlsx(wb, sheet = "HOME")
      proposer <- read.xlsx(wb, sheet = "1. PROPOSER")
      site <- read.xlsx(wb, sheet = "2. SITE")
      species <- read.xlsx(wb, sheet = "3. SPECIES")
      ecosystems <- read.xlsx(wb, sheet = "4. ECOSYSTEMS & C")
      threats <- read.xlsx(wb, sheet = "5. THREATS")
      review <- read.xlsx(wb, sheet = "6. REVIEW")
      citations <- read.xlsx(wb, sheet = "7. CITATIONS")
      check <- read.xlsx(wb, sheet = "8. CHECK")
      
            # Invisible sheets
      checkboxes <- read.xlsx(wb, sheet = "checkboxes")
      resultsSpecies <- read.xlsx(wb, sheet = "results_species")
      resultsEcosystems <- read.xlsx(wb, sheet = "results_ecosystems")
    
      # Handle the site name
            # Get the name
      name <- site[1,"GENERAL"]
      
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
      if(nchar(name)>255){
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
      
      # Form version number
            # Get version
      formVersion <- home[1,1] %>% substr(., start=9, stop=nchar(.)) %>% as.numeric()
      
            # Check compatibility
      if(!formVersion %in% c(1, 1.1, 1.2)){
        convert_res[step,"Result"] <- emo::ji("prohibited")
        convert_res[step,"Message"] <- "Form version not supported. Please contact Chloé and provide her with this error message."
        KBAforms[step] <- NA
        next
      }
      
      # Format the sheets
            # 1. PROPOSER
      proposer %<>%
        .[, 2:3] %>%
        rename(Field = X2, Entry = X3) %>%
        filter(!is.na(Field))
      
            # 2. SITE
                  # General
      site %<>%
        .[, 2:4] %>%
        rename(Field = X2) %>%
        filter(!Field == "Ongoing                                                                                           Needed                                                  ")
      
                  # Conservation actions
      actionsCol <- which(colnames(checkboxes) == "2..Conservation.actions")
      
      actions <- checkboxes %>%
        .[, actionsCol:(actionsCol+2)]
      colnames(actions) <- actions[1,]
      actions %<>%
        .[2:nrow(.),]
      
            # 3. SPECIES
                  # General
      colnames(species) <- species[1,]
      species %<>%
        .[2:nrow(.),] %>%
        filter(!is.na(`Common name`)) %>%
        mutate(`Common name` = trimws(`Common name`),
               `Scientific name` = trimws(`Scientific name`),
               Sensitive = F) %>%
        mutate(`Derivation of best estimate` = ifelse(`Derivation of best estimate` == "Other (please add further details in column AA)", "Other", `Derivation of best estimate`))
      
      if(language == "french"){
        species %<>%
          left_join(., xwalk_derivationOfBestEstimate, by=c("Derivation of best estimate" = "DerivationOfBestEstimate_EN")) %>%
          mutate(`Derivation of best estimate` = DerivationOfBestEstimate_FR) %>%
          select(-DerivationOfBestEstimate_FR)
      }
      
      if(formVersion %in% c(1, 1.1)){
        colnames(species)[which(colnames(species) == "RU Source")] <- "RU source"
      }
      
      # If French is requested, translate the species common names to French
      if(language == "french"){
        
        species %<>%
          left_join(., masterSpeciesList[,c("ELEMENT_CODE", "NATIONAL_FR_NAME")], by=c("NatureServe Element Code" = "ELEMENT_CODE")) %>%
          mutate(`Common name` = NATIONAL_FR_NAME) %>%
          select(-NATIONAL_FR_NAME)
        
        if(sum(is.na(species$`Common name`)) > 0){
          
          if(!sum(species$`NatureServe Element Code` %in% masterSpeciesList$ELEMENT_CODE) == nrow(species)){
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
      for(i in 1:nrow(species)){
        
        if(grepl(";", species$`Common name`[i])){
          species$`Common name`[i] %<>% substr(., start=1, stop=unlist(gregexpr(";", species$`Common name`[i]))-1)
        }
      }
      
      # Only retain information in the desired language
      error <- F
      
      for(i in 1:nrow(species)){
        
        for(j in 1:ncol(species)){
          
          if(grepl("FRANCAIS", species[i,j]) | grepl("ENGLISH", species[i,j])){
            
            # Initiate language check
            checkFR <- F
            checkEN <- F
            
            # Get index of FRANCAIS annotation
            if(grepl("FRANCAIS", species[i,j])){
              checkFR <- T
              startFR <- unlist(gregexpr("FRANCAIS", species[i,j]))
            }
            
            # Get index of ENGLISH annotation
            if(grepl("ENGLISH", species[i,j])){
              checkEN <- T
              startEN <- unlist(gregexpr("ENGLISH", species[i,j]))
            }
            
            # Get desired text
            if(checkFR & checkEN){
              
              if(startFR < startEN){
                FR <- substr(species[i,j], start=startFR + nchar("FRANCAIS"), stop=startEN-1)
                EN <- substr(species[i,j], start=startEN + nchar("ENGLISH"), stop=nchar(species[i,j]))
                  
              }else{
                FR <- substr(species[i,j], start=startFR + nchar("FRANCAIS"), stop=nchar(species[i,j]))
                EN <- substr(species[i,j], start=startEN + nchar("ENGLISH"), stop=startFR-1)
              }
              
              if(language == "english"){
                final <- EN
                
              }else{
                final <- FR
              }
              
            }else if(checkFR){
              
              if(language == "french"){
                final <- substr(species[i,j], start=startFR + nchar("FRANCAIS"), stop=nchar(species[i,j]))
                
              }else{
                convert_res[step,"Result"] <- emo::ji("prohibited")
                convert_res[step,"Message"] <- paste0("The summary was requested in English, but information in the '", colnames(species)[j], "' field (3. SPECIES tab) is not provided in English. Please enter the information in English, preceded by the text 'ENGLISH -'.")
                KBAforms[step] <- NA
                error <- T
                break
              }
              
            }else if(checkEN){
              
              if(language == "english"){
                final <- substr(species[i,j], start=startEN + nchar("ENGLISH"), stop=nchar(species[i,j]))
                
              }else{
                convert_res[step,"Result"] <- emo::ji("prohibited")
                convert_res[step,"Message"] <- paste0("The summary was requested in French, but information in the '", colnames(species)[j], "' field (3. SPECIES tab) is not provided in French. Please enter the information in French, preceded by the text 'FRANCAIS -'.")
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
            species[i,j] <- final
            
          }else{
            
            if(language=="french" & colnames(species)[j] %in% c("Composition of 10 RUs", "Explanation of site estimates", "Explanation of reference estimates") & !is.na(species[i,j])){
              convert_res[step,"Result"] <- emo::ji("prohibited")
              convert_res[step,"Message"] <- paste0("The summary was requested in French, but information in the '", colnames(species)[j], "' field (3. SPECIES tab) is not provided in French. Please enter the information in French, preceded by the text 'FRANCAIS -'. Information in English should be preceded by 'ENGLISH -'.")
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
      if((!formVersion %in% c(1, 1.1)) & (reviewStage == "general")){
        
            # Check that the Public Display section is filled out
        if(sum(is.na(species$`Display taxonomic group?`), is.na(species$`Display taxon name?`), is.na(species$`Display assessment information?`), is.na(species$`Display internal boundary?`)) > 0){
          convert_res[step,"Result"] <- emo::ji("prohibited")
          convert_res[step,"Message"] <- "You are requesting a summary for General Review and the Public Display section of the KBA Canada Proposal Form is not filled out. Please fill out this section before you proceed with General Review."
          KBAforms[step] <- NA
          next
          
        }else{
          
          for(i in 1:nrow(species)){
            
            alternativeName <- species$`Alternative name to display`[i] %>%
              str_to_sentence()
            if(language == "english"){
              alternativeName <- ifelse(is.na(alternativeName) || alternativeName == "", "A sensitive taxon", alternativeName)
            }else{
              alternativeName <- ifelse(is.na(alternativeName) || alternativeName == "", "Un taxon sensible", alternativeName)
            }
            
              # Display taxonomic group?
            if(species$`Display taxonomic group?`[i] == "No"){
              species$`Taxonomic group`[i] <- "-"
              species$`Common name`[i] <- alternativeName
              species$`Scientific name`[i] <- alternativeName
              species$Sensitive[i] <- T
            }
            
              # Display taxon name?
            if(species$`Display taxon name?`[i] == "No"){
              species$`Common name`[i] <- alternativeName
              species$`Scientific name`[i] <- alternativeName
              species$Sensitive[i] <- T
            }
            
              # Display assessment information?
            if(species$`Display assessment information?`[i] == "No"){
              species$Status[i] <- "-"
              species$`Status assessment agency`[i] <- "-"
              species$`Reproductive Units (RU)`[i] <- "-"
              species$`Assessment parameter`[i] <- "(i) -"
              species$`Min site estimate`[i] <- "-"
              species$`Best site estimate`[i] <- "-"
              species$`Max site estimate`[i] <- "-"
              species$`Year of site estimate`[i] <- "-"
              species$`Min reference estimate`[i] <- "-"
              species$`Best reference estimate`[i] <- "-"
              species$`Max reference estimate`[i] <- "-"
              species$`Composition of 10 RUs`[i] <- "-"
              species$`RU source`[i] <- "-"
              species$`Derivation of best estimate`[i] <- "-"
              species$`Explanation of site estimates`[i] <- "-"
              species$`Sources of site estimates`[i] <- "-"
              species$`Explanation of reference estimates`[i] <- "-"
              species$`Sources of reference estimates`[i] <- "-"
              species$Sensitive[i] <- T
            }
          }
        }
      }
      
      # Sort by scientific name
      species %<>% arrange(`Scientific name`)
      
            # 4. ECOSYSTEMS & C
      ecosystems %<>%
        pull(X2) %>%
        unique() %>%
        .[which(!. == "Criteria met")]
      
      if(!length(ecosystems) == 0){
        convert_res[step,"Result"] <- emo::ji("prohibited")
        convert_res[step,"Message"] <- "Ecosystem KBAs not yet supported. Please contact Chloé and provide her with this error message."
        KBAforms[step] <- NA
        next
      }
            
            # 5. THREATS
                  # Verify whether "No Threats" checkbox is checked
      noThreatsCol <- which(colnames(checkboxes) == "5..Threats")
      
      noThreats <- checkboxes[2, (noThreatsCol+1)] %>% as.logical()
      
                  # If there are threats, get that information
      if(!noThreats){
        colnames(threats) <- threats[3,]
        threats %<>% .[4:nrow(.),]
        colnames(threats)[ncol(threats)] <- "Notes"
      }  
         
            # 6. REVIEW
                  # General
      review %<>%
        drop_na(X2) %>%
        fill(`INSTRUCTIONS:`)
    
                  # Technical review
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
      
                  # General review
      generalReview <- review %>%
        filter(`INSTRUCTIONS:` == 2) %>%
        select(-c(`INSTRUCTIONS:`))
      
      if(is.na(generalReview[2,4])){
        generalReview %<>% select(-X5)
      }
      
      colnames(generalReview) <- generalReview[2,]
      
      if(nrow(generalReview) > 2){
        generalReview %<>% .[3:nrow(.),]
      }else{
        generalReview[3,] <- c("No reviewers listed", rep("", (ncol(generalReview)-1)))
        generalReview %<>% .[3:nrow(.),]
      }
       
            # 7. CITATIONS
                  # General
      colnames(citations) <- tolower(citations[2,])
      citations %<>%
        .[3:nrow(.), 1:4] %>%
        filter(!is.na(`short citation`))
      
                  # Redact sensitive citations
      if(reviewStage == "general"){
        citations %<>% mutate(Sensitive = ifelse(grepl("SENSITIVE", citations$`short citation`, T), T, F)) %>%
          filter(!Sensitive)
      }
        
            # 8. CHECK
                  # Column names
      colnames(check) <- c("Check", "Item")
      
                  # Get checkbox results
      check_checkboxes <- checkboxes %>%
        .[2:nrow(.),] %>%
        select("8..Checks") %>%
        drop_na()
      
      if(formVersion %in% c(1, 1.1)){
        check_checkboxes %<>% .[c(1:5,7:nrow(.)),] # Cell N8 is obsolete in v1.1 of the Proposal Form (it doens't link to any actual checkbox)
      }else{
        check_checkboxes %<>% pull(`8..Checks`)
      }
          
                  # Verify that there are as many checkbox results as there are checkboxes
      if(!(nrow(check) == length(check_checkboxes))){
        convert_res[step,"Result"] <- emo::ji("prohibited")
        convert_res[step,"Message"] <- "Inconsistencies between the 8. CHECKS tab and checkbox results. This error originates from the Excel formulas themselves. Please contact Chloé and provide her with this error message."
        KBAforms[step] <- NA
        next
      }
      
                  # Add checkbox results to the 8. CHECK tab
      check %<>%
        select(-Check) %>%
        mutate(Check = check_checkboxes)
      rm(check_checkboxes)
      
      # Prepare variables
            # 1. KBA Name
      nationalName <<- site$GENERAL[which(site$Field == "National name")]
      
            # 2. Location
                  # Jurisdiction
      juris <<- site$GENERAL[which(site$Field == "Province or Territory")]
      
      if(language == "french"){
        juris <<- xwalk_jurisdiction %>%
          .[which(.$Province_EN == juris), "Province_FR"]
      }
      
                  # Latitude and Longitude
      lat <<- site$GENERAL[which(site$Field == "Latitude (dd.dddd)")] %>%
        as.numeric(.) %>%
        round(., 3)
      if(language == "english"){
        lat <<- ifelse(is.na(lat), "coordinates unspecified", lat)
      }else{
        lat <<- ifelse(is.na(lat), "coordonnées non spécifiées", lat)
      }
      
      lon <<- site$GENERAL[which((site$Field == "Longitude (dd.dddd)" | site$Field == "Longitude (ddd.dddd)"))] %>%
        as.numeric(.) %>%
        round(., 3)
      lon <<- ifelse(is.na(lon), "", paste0("/", lon))
      
            # 3. KBA Scope
      criteriaMet <- home$X4[which(home$X3 == "Criteria met")]
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
      if(formVersion %in% c(1, 1.1)){
        proposalLead <<- proposer$Entry[which(proposer$Field == "Name")]
      }else{
        proposalLead <<- proposer$Entry[which(proposer$Field == "Name of proposal development lead")]
      }
      
            # 7. Site Description
      if(language == "english"){
        siteDescription <<- site$GENERAL[which(site$Field == "Site description")]
      }else{
        siteDescription <<- site$FRENCH[which(site$Field == "Site description")]
      }
      
            # 8. Assessment Details - KBA Trigger Species
      if(language == "english"){
        includeGlobalTriggers <<- ifelse(scope %in% c("Global and National", "Global"), "GLOBAL", "")
        includeNationalTriggers <<- ifelse(scope %in% c("Global and National", "National"), "NATIONAL", "")
      }else{
        includeGlobalTriggers <<- ifelse(scope %in% c("Mondial et National", "Mondial"), "NIVEAU MONDIAL", "")
        includeNationalTriggers <<- ifelse(scope %in% c("Mondial et National", "National"), "NIVEAU NATIONAL", "")
      }
      
            # 10. Delineation Rationale
      if(language == "english"){
        delineationRationale <<- site$GENERAL[which(site$Field == "Delineation rationale")]
      }else{
        delineationRationale <<- site$FRENCH[which(site$Field == "Delineation rationale")]
      }
      
            # 12. General Review
      noFeedback <<- review$X3[which(review$X2 == "Provide information about any organizations you contacted and that did not provide feedback.")]
      noFeedback <<- ifelse(is.na(noFeedback), "None", noFeedback)
      
            # 13. Additional Site Information
      if(language == "english"){
        nominationRationale <- site$GENERAL[which(site$Field == "Rationale for nomination")]
        additionalBiodiversity <- site$GENERAL[which(site$Field == "Additional biodiversity")]
        customaryJurisdiction <- site$GENERAL[which(site$Field == "Customary jurisdiction")]
      }else{
        nominationRationale <- site$FRENCH[which(site$Field == "Rationale for nomination")]
        additionalBiodiversity <- site$FRENCH[which(site$Field == "Additional biodiversity")]
        customaryJurisdiction <- site$FRENCH[which(site$Field == "Customary jurisdiction")]
      }
      
      if(formVersion %in% c(1, 1.1)){
        
        siteHistory <- NA
        
        if(language == "english"){
          conservation <- site$GENERAL[which(site$Field == "Site management")]
        }else{
          conservation <- site$FRENCH[which(site$Field == "Site management")]
        }
        
      }else{
        
        if(language == "english"){
          customaryJurisdictionSource <- site$GENERAL[which(site$Field == "Customary jurisdiction source")]
          siteHistory <- site$GENERAL[which(site$Field == "Site history")]
          conservation <- site$GENERAL[which(site$Field == "Conservation")]
        }else{
          customaryJurisdictionSource <- site$FRENCH[which(site$Field == "Customary jurisdiction source")]
          siteHistory <- site$FRENCH[which(site$Field == "Site history")]
          conservation <- site$FRENCH[which(site$Field == "Conservation")]
        }
      }
      
      # Prepare flextables
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
      
                  # Number of species
      maxCol <- max(sapply(species$`Criteria met`[which(!is.na(species$`Criteria met`))], function(x) str_count(x, ";")))+1
      criteriaCols <- paste0("Col", 1:maxCol)
      
      criteriaInfo <- species %>%
        filter(!is.na(`Criteria met`)) %>%
        select(`Scientific name`, `Criteria met`) %>%
        separate(`Criteria met`, into=criteriaCols, sep="; ", fill="right") %>%
        pivot_longer(all_of(criteriaCols), names_to = "Remove", values_to="Criteria met") %>%
        filter(!is.na(`Criteria met`)) %>%
        group_by(`Criteria met`) %>%
        summarise(NSpecies = n(), .groups="drop") %>%
        left_join(criteriaInfo, ., by=c("CriteriaFull" = "Criteria met"))
      
                  # Species names
      criteriaInfo <- species %>%
        filter(!is.na(`Criteria met`)) %>%
        select(`Scientific name`, `Criteria met`) %>%
        separate(`Criteria met`, into=criteriaCols, sep="; ", fill="right") %>%
        pivot_longer(all_of(criteriaCols), names_to = "Remove", values_to="Criteria met") %>%
        filter(!is.na(`Criteria met`)) %>%
        arrange(`Scientific name`) %>%
        group_by(`Criteria met`) %>%
        summarise(speciesNames = paste(`Scientific name`, collapse=", "), .groups="drop") %>%
        left_join(criteriaInfo, ., by=c("CriteriaFull" = "Criteria met"))
      
                  # Flextable
      criteriaInfo_ft <- criteriaInfo %>%
        mutate(Label = "") %>%
        mutate(Blank = "") %>%
        flextable(col_keys = c("Blank", "Label"))
      
      if(language == "english"){
        criteriaInfo_ft %<>% compose(j='Label', value=as_paragraph(as_chunk(x=paste0(as.character("\u25CF"), " ", Scope, " ", Criteria, " [criterion met by ", NSpecies, ifelse(NSpecies == 1, " taxon]", " taxa]"), " - ", Definition, " (")), as_chunk(x=speciesNames, props=fp_text(font.size=11, font.family='Calibri', italic=T)), as_chunk(x=").")))
      }else{
        criteriaInfo_ft %<>% compose(j='Label', value=as_paragraph(as_chunk(x=paste0(as.character("\u25CF"), " ", Criteria, " ", Scope, " [critère rempli par ", NSpecies, ifelse(NSpecies == 1, " taxon]", " taxons]"), " - ", Definition, " (")), as_chunk(x=speciesNames, props=fp_text(font.size=11, font.family='Calibri', italic=T)), as_chunk(x=").")))
      }
       
      criteriaInfo_ft %<>% 
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
        select(`Scientific name`, Status, `Criteria met`, `Reproductive Units (RU)`, `Composition of 10 RUs`, `RU source`, AssessmentParameter, Blank, SiteEstimate_Min, SiteEstimate_Best, SiteEstimate_Max, `Year of site estimate`, `Derivation of best estimate`, `Explanation of site estimates`, `Sources of site estimates`, TotalEstimate_Min, TotalEstimate_Best, TotalEstimate_Max, `Explanation of reference estimates`, `Sources of reference estimates`, PercentAtSite, Sensitive)
      
                  # Separate global and national assessments
      speciesAssessments_g <- speciesAssessments %>%
        filter(grepl("g", `Criteria met`, fixed=T)) %>%
        mutate(`Criteria met` = substr(`Criteria met`, start=2, stop=nchar(`Criteria met`)))
      
      speciesAssessments_n <- speciesAssessments %>%
        filter(grepl("n", `Criteria met`, fixed=T)) %>%
        mutate(`Criteria met` = substr(`Criteria met`, start=2, stop=nchar(`Criteria met`)))
      
      if(!(nrow(speciesAssessments_g) + nrow(speciesAssessments_n)) == nrow(speciesAssessments)){
        convert_res[step,"Result"] <- emo::ji("prohibited")
        convert_res[step,"Message"] <- "Some assessments are not being correctly classified as global or national assessments. This is an error with the code. Please contact Chloé and provide her with this error message."
        KBAforms[step] <- NA
        next
    	}
      rm(speciesAssessments)
      
                  # Information for the footnotes
      footnotes_g <- speciesAssessments_g %>%
        select(`Composition of 10 RUs`, `RU source`, `Derivation of best estimate`, `Explanation of site estimates`, `Sources of site estimates`, `Explanation of reference estimates`, `Sources of reference estimates`, Sensitive) %>%
        mutate(`Composition of 10 RUs` = sapply(`Composition of 10 RUs`, function(x) ifelse(substr(x, start=nchar(x), stop=nchar(x)) == ".", x, paste0(x, ".")))) %>%
        mutate(`RU source` = sapply(`RU source`, function(x) ifelse(substr(x, start=nchar(x), stop=nchar(x)) == ".", x, paste0(x, ".")))) %>%
        mutate(`Derivation of best estimate` = sapply(`Derivation of best estimate`, function(x) ifelse(substr(x, start=nchar(x), stop=nchar(x)) == ".", x, paste0(x, ".")))) %>% 
        mutate(`Explanation of site estimates` = sapply(`Explanation of site estimates`, function(x) ifelse(substr(x, start=nchar(x), stop=nchar(x)) == ".", x, paste0(x, ".")))) %>%
        mutate(`Sources of site estimates` = sapply(`Sources of site estimates`, function(x) ifelse(substr(x, start=nchar(x), stop=nchar(x)) == ".", x, paste0(x, ".")))) %>%
        mutate(`Explanation of reference estimates` = sapply(`Explanation of reference estimates`, function(x) ifelse(substr(x, start=nchar(x), stop=nchar(x)) == ".", x, paste0(x, ".")))) %>%
        mutate(`Sources of reference estimates` = sapply(`Sources of reference estimates`, function(x) ifelse(substr(x, start=nchar(x), stop=nchar(x)) == ".", x, paste0(x, "."))))
      
      if(language == "english"){
        footnotes_g %<>%
          mutate(RU_Source = ifelse(is.na(`Composition of 10 RUs`) & is.na(`RU source`), NA, paste0("Composition of 10 Reproductive Units (RUs): ", `Composition of 10 RUs`, " Source of RU data: ", `RU source`))) %>%
          mutate(Site_Source = paste0("Derivation of site estimate: ", `Derivation of best estimate`, " Explanation of site estimate(s): ", `Explanation of site estimates`, " Source(s) of site estimate(s): ", `Sources of site estimates`)) %>%
          mutate(Reference_Source = paste0("Explanation of global estimate(s): ", `Explanation of reference estimates`, " Source(s) of global estimate(s): ", `Sources of reference estimates`)) %>%
          select(RU_Source, Site_Source, Reference_Source, Sensitive)
      }else{
        footnotes_g %<>%
          mutate(RU_Source = ifelse(is.na(`Composition of 10 RUs`) & is.na(`RU source`), NA, paste0("Composition de 10 Unités Reproductives (URs) : ", `Composition of 10 RUs`, " Source des données d'URs : ", `RU source`))) %>%
          mutate(Site_Source = paste0("Calcul de l'estimation au site : ", `Derivation of best estimate`, " Explication de(s) estimation(s) au site : ", `Explanation of site estimates`, " Source(s) de(s) estimation(s) au site : ", `Sources of site estimates`)) %>%
          mutate(Reference_Source = paste0("Explication de(s) estimation(s) mondiale(s) : ", `Explanation of reference estimates`, " Source(s) de(s) estimation(s) mondiale(s) : ", `Sources of reference estimates`)) %>%
          select(RU_Source, Site_Source, Reference_Source, Sensitive)
      }
        
      footnotes_n <- speciesAssessments_n %>%
        select(`Composition of 10 RUs`, `RU source`, `Derivation of best estimate`, `Explanation of site estimates`, `Sources of site estimates`, `Explanation of reference estimates`, `Sources of reference estimates`, Sensitive) %>%
        mutate(`Composition of 10 RUs` = sapply(`Composition of 10 RUs`, function(x) ifelse(substr(x, start=nchar(x), stop=nchar(x)) == ".", x, paste0(x, ".")))) %>%
        mutate(`RU source` = sapply(`RU source`, function(x) ifelse(substr(x, start=nchar(x), stop=nchar(x)) == ".", x, paste0(x, ".")))) %>%
        mutate(`Derivation of best estimate` = sapply(`Derivation of best estimate`, function(x) ifelse(substr(x, start=nchar(x), stop=nchar(x)) == ".", x, paste0(x, ".")))) %>% 
        mutate(`Explanation of site estimates` = sapply(`Explanation of site estimates`, function(x) ifelse(substr(x, start=nchar(x), stop=nchar(x)) == ".", x, paste0(x, ".")))) %>%
        mutate(`Sources of site estimates` = sapply(`Sources of site estimates`, function(x) ifelse(substr(x, start=nchar(x), stop=nchar(x)) == ".", x, paste0(x, ".")))) %>%
        mutate(`Explanation of reference estimates` = sapply(`Explanation of reference estimates`, function(x) ifelse(substr(x, start=nchar(x), stop=nchar(x)) == ".", x, paste0(x, ".")))) %>%
        mutate(`Sources of reference estimates` = sapply(`Sources of reference estimates`, function(x) ifelse(substr(x, start=nchar(x), stop=nchar(x)) == ".", x, paste0(x, "."))))
      
      if(language == "english"){
        footnotes_n %<>%
          mutate(RU_Source = ifelse(is.na(`Composition of 10 RUs`) & is.na(`RU source`), NA, paste0("Composition of 10 Reproductive Units (RUs): ", `Composition of 10 RUs`, " Source of RU data: ", `RU source`))) %>%
          mutate(Site_Source = paste0("Derivation of site estimate: ", `Derivation of best estimate`, " Explanation of site estimate(s): ", `Explanation of site estimates`, " Source(s) of site estimate(s): ", `Sources of site estimates`)) %>%
          mutate(Reference_Source = paste0("Explanation of national estimate(s): ", `Explanation of reference estimates`, " Source(s) of national estimate(s): ", `Sources of reference estimates`)) %>%
          select(RU_Source, Site_Source, Reference_Source, Sensitive)
      }else{
        footnotes_n %<>%
          mutate(RU_Source = ifelse(is.na(`Composition of 10 RUs`) & is.na(`RU source`), NA, paste0("Composition de 10 Unités Reproductives (URs) : ", `Composition of 10 RUs`, " Source des données d'URs : ", `RU source`))) %>%
          mutate(Site_Source = paste0("Calcul de l'estimation au site : ", `Derivation of best estimate`, " Explication de(s) estimation(s) au site : ", `Explanation of site estimates`, " Source(s) de(s) estimation(s) au site : ", `Sources of site estimates`)) %>%
          mutate(Reference_Source = paste0("Explication de(s) estimation(s) nationale(s) : ", `Explanation of reference estimates`, " Source(s) de(s) estimation(s) nationale(s) : ", `Sources of reference estimates`)) %>%
          select(RU_Source, Site_Source, Reference_Source, Sensitive)
      }
      
                  # Information for the main table
      speciesAssessments_g %<>% select(-c(`Composition of 10 RUs`, `RU source`, `Derivation of best estimate`, `Explanation of site estimates`, `Sources of site estimates`, `Explanation of reference estimates`, `Sources of reference estimates`))
      
      speciesAssessments_n %<>% select(-c(`Composition of 10 RUs`, `RU source`, `Derivation of best estimate`, `Explanation of site estimates`, `Sources of site estimates`, `Explanation of reference estimates`, `Sources of reference estimates`))
      
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
            select(-Sensitive) %>%
            flextable()
          
          if(language == "english"){
            speciesAssessments_g_ft %<>%
              width(j=colnames(.), width=c(1.5,1.3,0.65,1.3,1.3,0.05,0.7,0.7,0.7,0.8)) %>%
              set_header_labels(values=list(`Scientific name` = "Species", Status = "Status*", `Criteria met`="Criteria Met", `Reproductive Units (RU)` = "# of Reproductive Units", AssessmentParameter = 'Assessment Parameter', Blank='', SiteEstimate_Best = "Value", `Year of site estimate` = "Year", TotalEstimate_Best = 'Global Estimate', PercentAtSite = "% of Global Pop. at Site")) %>%
              add_header_row(values = c("Species", "Status*", "Criteria Met", "# of Reproductive Units", "Assessment Parameter", "", "Site Estimate", 'Global Estimate', "% of Global Pop. at Site"), colwidths=c(1, 1, 1, 1, 1, 1, 2, 1, 1))
          }else{
            speciesAssessments_g_ft %<>%
              width(j=colnames(.), width=c(1.5,1,0.85,1.2,1.3,0.05,0.7,0.7,0.9,0.8)) %>%
              set_header_labels(values=list(`Scientific name` = "Espèce", Status = "Statut*", `Criteria met`="Critère(s) atteint(s)", `Reproductive Units (RU)` = "# d’Unités Reproductives", AssessmentParameter = 'Paramètre d’évaluation', Blank='', SiteEstimate_Best = "Valeur", `Year of site estimate` = "Année", TotalEstimate_Best = 'Estimation mondiale', PercentAtSite = "% de la pop. mondiale au site")) %>%
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
            select(-Sensitive) %>%
            mutate(Blank2 = "") %>%
            relocate(Blank2, .after = `Year of site estimate`) %>%
            flextable()
          
          if(language == "english"){
            speciesAssessments_g_ft %<>%
              width(j=colnames(.), width=c(1.4,1.2,0.65,1.1,0.9,0.05,0.4,0.4,0.4,0.5,0.05,0.4,0.4,0.4,0.8)) %>%
              set_header_labels(values=list(`Scientific name` = "Species", Status = "Status*", `Criteria met`="Criteria Met", `Reproductive Units (RU)` = "# of Reproductive Units", AssessmentParameter = 'Assessment Parameter', Blank='', SiteEstimate_Min = "Min", SiteEstimate_Best = "Best", SiteEstimate_Max = "Max", SiteEstimate_Year = "Year", Blank2 = "", TotalEstimate_Min = "Min", TotalEstimate_Best = "Best", TotalEstimate_Max = "Max", PercentAtSite = "% of Global Pop. at Site")) %>%
              add_header_row(values = c("Species", "Status*", "Criteria Met", "# of Reproductive Units", "Assessment Parameter", "", "Site Estimate", "", "Global Estimate", "% of Global Pop. at Site"), colwidths=c(1, 1, 1, 1, 1, 1, 4, 1, 3, 1))
          }else{
            speciesAssessments_g_ft %<>%
              width(j=colnames(.), width=c(0.9,0.8,0.75,0.8,1,0.05,0.4,0.8,0.5,0.6,0.05,0.4,0.8,0.5,0.8)) %>%
              set_header_labels(values=list(`Scientific name` = "Espèce", Status = "Statut*", `Criteria met`="Critère(s) atteint(s)", `Reproductive Units (RU)` = "# d’Unités Reprod.", AssessmentParameter = 'Paramètre d’évaluation', Blank='', SiteEstimate_Min = "Min", SiteEstimate_Best = "Meilleure", SiteEstimate_Max = "Max", `Year of site estimate` = "Année", TotalEstimate_Min = "Min", TotalEstimate_Best = 'Meilleure', TotalEstimate_Max = "Max", PercentAtSite = "% de la pop. mondiale au site")) %>%
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
            select(-Sensitive) %>%
            flextable()
          
          if(language == "english"){
            speciesAssessments_n_ft %<>%
              width(j=colnames(.), width=c(1.5,1.3,0.65,1.3,1.3,0.05,0.7,0.7,0.7,0.8)) %>%
              set_header_labels(values=list(`Scientific name` = "Taxon", Status = "Status*", `Criteria met`="Criteria Met", `Reproductive Units (RU)` = "# of Reproductive Units", AssessmentParameter = 'Assessment Parameter', Blank='', SiteEstimate_Best = "Value", `Year of site estimate` = "Year", TotalEstimate_Best = 'National Estimate', PercentAtSite = "% of National Pop. at Site")) %>%
              add_header_row(values = c("Taxon", "Status*", "Criteria Met", "# of Reproductive Units", "Assessment Parameter", "", "Site Estimate", 'National Estimate', "% of National Pop. at Site"), colwidths=c(1, 1, 1, 1, 1, 1, 2, 1, 1))
          }else{
            speciesAssessments_n_ft %<>%
              width(j=colnames(.), width=c(1.5,1,0.85,1.2,1.3,0.05,0.7,0.7,0.9,0.8)) %>%
              set_header_labels(values=list(`Scientific name` = "Taxon", Status = "Statut*", `Criteria met`="Critère(s) atteint(s)", `Reproductive Units (RU)` = "# d’Unités Reproductives", AssessmentParameter = 'Paramètre d’évaluation', Blank='', SiteEstimate_Best = "Valeur", `Year of site estimate` = "Année", TotalEstimate_Best = 'Estimation nationale', PercentAtSite = "% de la pop. nationale au site")) %>%
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
            select(-Sensitive) %>%
            mutate(Blank2 = "") %>%
            relocate(Blank2, .after = `Year of site estimate`) %>%
            flextable()
          
          if(language == "english"){
            speciesAssessments_n_ft %<>%
              width(j=colnames(.), width=c(1.4,1.2,0.65,1.1,0.9,0.05,0.4,0.4,0.4,0.5,0.05,0.4,0.4,0.4,0.8)) %>%
              set_header_labels(values=list(`Scientific name` = "Taxon", Status = "Status*", `Criteria met`="Criteria Met", `Reproductive Units (RU)` = "# of Reproductive Units", AssessmentParameter = 'Assessment Parameter', Blank='', SiteEstimate_Min = "Min", SiteEstimate_Best = "Best", SiteEstimate_Max = "Max", SiteEstimate_Year = "Year", Blank2 = "", TotalEstimate_Min = "Min", TotalEstimate_Best = "Best", TotalEstimate_Max = "Max", PercentAtSite = "% of National Pop. at Site")) %>%
              add_header_row(values = c("Taxon", "Status*", "Criteria Met", "# of Reproductive Units", "Assessment Parameter", "", "Site Estimate", "", "National Estimate", "% of National Pop. at Site"), colwidths=c(1, 1, 1, 1, 1, 1, 4, 1, 3, 1))
          }else{
            speciesAssessments_n_ft %<>%
              width(j=colnames(.), width=c(0.9,0.8,0.75,0.8,1,0.05,0.4,0.8,0.5,0.6,0.05,0.4,0.8,0.5,0.8)) %>%
              set_header_labels(values=list(`Scientific name` = "Taxon", Status = "Statut*", `Criteria met`="Critère(s) atteint(s)", `Reproductive Units (RU)` = "# d’Unités Reprod.", AssessmentParameter = 'Paramètre d’évaluation', Blank='', SiteEstimate_Min = "Min", SiteEstimate_Best = "Meilleure", SiteEstimate_Max = "Max", `Year of site estimate` = "Année", Blank2 = "", TotalEstimate_Min = "Min", TotalEstimate_Best = 'Meilleure', TotalEstimate_Max = "Max", PercentAtSite = "% de la pop. nationale au site")) %>%
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
          
          if(!speciesAssessments_g$Sensitive[i]){
            
            for(c in 1:ncol(footnotes_g %>% select(-Sensitive))){
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
          }else{
            footnote <- footnote+1
            
            if(language == "english"){
              speciesAssessments_g_ft %<>% footnote(i=i, j=1, value=as_paragraph(as.character("For more information, please contact the KBA Canada Secretariat.")), ref_symbols=as.integer(footnote))
            }else{
              speciesAssessments_g_ft %<>% footnote(i=i, j=1, value=as_paragraph(as.character("Pour d'avantage d'informations, merci de contacter le Secrétariat KBA Canada.")), ref_symbols=as.integer(footnote))
            }
          }
        }
      }
     
            # National
      if(nrow(speciesAssessments_n) > 0){
        footnote <- 0
        for(i in 1:nrow(speciesAssessments_n)){
          col <- which(grepl("http", footnotes_n[i,]), arr.ind = TRUE)
          
          if(!speciesAssessments_n$Sensitive[i]){
            
            for(c in 1:ncol(footnotes_n %>% select(-Sensitive))){
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
          }else{
            footnote <- footnote+1
            speciesAssessments_n_ft %<>% footnote(i=i, j=1, value=as_paragraph(as.character("For more information, please contact the KBA Canada Secretariat.")), ref_symbols=as.integer(footnote))
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
        mutate(Prefix = ifelse(language=="english",
                               paste0(as.character("\u25CF"), " Species: "),
                               paste0(as.character("\u25CF"), " Espèce(s) : ")))
      elementsSummary <- elementsSummary[,c(ncol(elementsSummary), 1:(ncol(elementsSummary)-1))]
      
      elementsSummary_ft <- flextable(elementsSummary, col_keys = c("Blank", "Label"), defaults=list(fontname="Calibri", font.size=11)) %>%
        width(j=colnames(.), width=c(0.3, 9))
      
      extraCall <- ""
      if(ncol(elementsSummary) > 3){
        
        # Keep only columns with common names
        spp <- 4:ncol(elementsSummary)
        spp <- spp[lapply(spp, "%%", 2) == 0]
        
        for(i in spp){
          
          if(!i == spp[length(spp)]){
            
            if(elementsSummary[i] == elementsSummary[i+1]){
              extraCall <- paste0(extraCall, ", as_chunk(x=', '), as_chunk(x=X", i-1, ")")
            }else{
              extraCall <- paste0(extraCall, ", as_chunk(x=', '), as_chunk(x=X", i-1, "), as_chunk(x=' ('), as_chunk(x=X", i, ", props=fp_text(font.size=11, font.family='Calibri', italic = T)), as_chunk(x=')')")
            }
            
          }else{
            
            if(elementsSummary[i] == elementsSummary[i+1]){
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
      
      if(elementsSummary[2] == elementsSummary[3]){
        compose_call <- paste0("elementsSummary_ft %<>% compose(j='Label', value=as_paragraph(as_chunk(Prefix), as_chunk(x=X1)", 
                               extraCall,
                               "))")
      }else{
        compose_call <- paste0("elementsSummary_ft %<>% compose(j='Label', value=as_paragraph(as_chunk(Prefix), as_chunk(x=X1), as_chunk(x=' ('), as_chunk(x=X2, props=fp_text(font.size=11, font.family='Calibri', italic = T)), as_chunk(x=')')", 
                               extraCall,
                               "))")
      }
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
      if(ncol(elementsSummary) > 2){
        
        # Keep only columns with scientific names
        spp <- 4:ncol(elementsSummary)
        spp <- spp[lapply(spp, "%%", 2) == 0]
        
        for(i in spp){
          
          if(!i == spp[length(spp)]){
            
            if(elementsSummary[i-1] == elementsSummary[i]){
              extraCall <- paste0(extraCall, ", as_chunk(x=', ', props=fp_text(font.size=12, font.family='Calibri', color='#5A5A5A')), as_chunk(x=X", i-1, ", props=fp_text(font.size=12, font.family='Calibri', color='#5A5A5A'))")
            }else{
              extraCall <- paste0(extraCall, ", as_chunk(x=', ', props=fp_text(font.size=12, font.family='Calibri', color='#5A5A5A')), as_chunk(x=X", i-1, ", props=fp_text(font.size=12, font.family='Calibri', color='#5A5A5A')), as_chunk(x=' (', props=fp_text(font.size=12, font.family='Calibri', color='#5A5A5A')), as_chunk(x=X", i, ", props=fp_text(font.size=12, font.family='Calibri', italic = T, color='#5A5A5A')), as_chunk(x=')', props=fp_text(font.size=12, font.family='Calibri', color='#5A5A5A'))")
            }
            
          }else{
            
            if(elementsSummary[i-1] == elementsSummary[i]){
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
      
      if(elementsSummary[1] == elementsSummary[2]){
        compose_call <- paste0("subtitle_ft %<>% compose(j='Label', value=as_paragraph(as_chunk(x=X1, props=fp_text(font.size=12, font.family='Calibri', color='#5A5A5A'))", 
                               extraCall,
                               "))")
      }else{
        compose_call <- paste0("subtitle_ft %<>% compose(j='Label', value=as_paragraph(as_chunk(x=X1, props=fp_text(font.size=12, font.family='Calibri', color='#5A5A5A')), as_chunk(x=' (', props=fp_text(font.size=12, font.family='Calibri', color='#5A5A5A')), as_chunk(x=X2, props=fp_text(font.size=12, font.family='Calibri', italic = T, color='#5A5A5A')), as_chunk(x=')', props=fp_text(font.size=12, font.family='Calibri', color='#5A5A5A'))", 
                               extraCall,
                               "))")
      }
      eval(parse(text=compose_call))
      
      subtitle_ft %<>%
        delete_part(part='header') %>%
        border_remove() %>%
        align(j=1, align = "left", part = "body") %>%
        fontsize(size=12, part='body')
      
      # Technical Review
      if(reviewStage == "general"){
        technicalReview_ft <- technicalReview %>%
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
      }
      
      # General Review
      generalReview_ft <- generalReview %>%
        .[,1:3] %>%
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
        speciesNotTriggers <- species %>%
          filter(is.na(`Criteria met`)) %>%
          pull(`Scientific name`) %>%
          unique() %>%
          paste(., collapse=", ")
      
        if(language == "english"){
          additionalInfo[nrow(additionalInfo)+1, ] <- c("Biodiversity elements that were assessed but did not meet KBA criteria", ifelse(speciesNotTriggers == "", "-", speciesNotTriggers))
        }else{
          additionalInfo[nrow(additionalInfo)+1, ] <- c("Éléments de biodiversité évalués qui n’atteignent pas les critères KBA", ifelse(speciesNotTriggers == "", "-", speciesNotTriggers))
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
      if(!formVersion %in% c(1, 1.1)){
        
        if(language == "english"){
          additionalInfo[nrow(additionalInfo)+1, ] <- c("Source of customary jurisdiction information", ifelse(is.na(customaryJurisdictionSource), "-", customaryJurisdictionSource))
        }else{
          additionalInfo[nrow(additionalInfo)+1, ] <- c("Source de l'information sur la juridiction coutumière", ifelse(is.na(customaryJurisdictionSource), "-", customaryJurisdictionSource))
        }
      }
      
            # Conservation
      additionalInfo[nrow(additionalInfo)+1, ] <- c("Conservation", ifelse(is.na(conservation), "-", conservation))
      
            # Ongoing conservation actions
      ongoingActions <- actions %>%
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
      if(!noThreats){
        threatText <- threats %>%
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
      neededActions <- actions %>%
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
      additionalInfo_ft <- additionalInfo %>%
        flextable() %>%
        width(j=colnames(.), width=c(3,6)) %>%
        font(fontname="Calibri", part="body") %>%
        fontsize(size=11, part='body')
      
      if(!reviewStage == "general"){
        additionalInfo_ft %<>% italic(i=2, j=2, part='body')
      }
      
      additionalInfo_ft %<>%
        delete_part(part="header") %>%
        theme_zebra() %>%
        align(j=c(1,2), align='left', part='body') %>%
        bold(j=1)
      
      # Citations
      if(nrow(citations) == 0){
        citations[1,] <- c("", "No references provided.", "", "")
      }
      
      citations_ft <- citations %>%
        arrange(`long citation`) %>%
        select(`long citation`) %>%
        flextable() %>%
        delete_part(part = "header") %>%
        width(j=colnames(.), width=9) %>%
        border_remove() %>%
        font(fontname="Calibri", part="body")
      
      # List all flextables
      if(scope %in% c("Global and National", "Mondial et National")){
        FT <- list(subtitle = subtitle_ft, triggerElements = elementsSummary_ft, criteria = criteriaInfo_ft, elements_g = elementsOnly_g, elementFootnotes_g = footnotesOnly_g, elements_n = elementsOnly_n, elementFootnotes_n = footnotesOnly_n, technicalReview = technicalReview_ft, generalReview = generalReview_ft, additionalInfo = additionalInfo_ft, references = citations_ft)
      }else if(scope %in% c("Global", "Mondial")){
        FT <- list(subtitle = subtitle_ft, triggerElements = elementsSummary_ft, criteria = criteriaInfo_ft, elements_g = elementsOnly_g, elementFootnotes_g = footnotesOnly_g, technicalReview = technicalReview_ft, generalReview = generalReview_ft, additionalInfo = additionalInfo_ft, references = citations_ft)
      }else{
        FT <- list(subtitle = subtitle_ft, triggerElements = elementsSummary_ft, criteria = criteriaInfo_ft, elements_n = elementsOnly_n, elementFootnotes_n = footnotesOnly_n, technicalReview = technicalReview_ft, generalReview = generalReview_ft, additionalInfo = additionalInfo_ft, references = citations_ft)
      }
      
      #### Save KBA summary
      # Get template
      if(language == "english"){
        if(reviewStage == "technical"){
          if(scope == "Global and National"){
            googledrive::drive_download("https://docs.google.com/document/d/1--Qh4Dif9Cr8RNS9u1ODcsVvXEDLBIEG", overwrite = TRUE)
            template <- "KBASummary_Template_NewForm_NoQuestions_TechnicalReview_GlobalNational.docx"
          }else if(scope == "Global"){
            googledrive::drive_download("https://docs.google.com/document/d/1-31LLlC09UpJeH6fKFFLagPtZG8jxkzT", overwrite = TRUE)
            template <- "KBASummary_Template_NewForm_NoQuestions_TechnicalReview_Global.docx"
          }else if(scope == "National"){
            googledrive::drive_download("https://docs.google.com/document/d/1mjDJVcLVkYGpc961QApZNU7YvuN4RqJc", overwrite = TRUE)
            template <- "KBASummary_Template_NewForm_NoQuestions_TechnicalReview_National.docx"
          }
        }else if(reviewStage == "general"){
          if(scope == "Global and National"){
            googledrive::drive_download("https://docs.google.com/document/d/1BPMhQJrxj_YksluSo1Nz06KkbJps703U", overwrite = TRUE)
            template <- "KBASummary_Template_NewForm_NoQuestions_GeneralReview_GlobalNational.docx"
          }else if(scope == "Global"){
            googledrive::drive_download("https://docs.google.com/document/d/1BJG4Sn71gl79grs2gjr7jNnE1WPr8UjZ", overwrite = TRUE)
            template <- "KBASummary_Template_NewForm_NoQuestions_GeneralReview_Global.docx"
          }else if(scope == "National"){
            googledrive::drive_download("https://docs.google.com/document/d/1BPwOfiReTd4Za5-c6zDY6kt2lnjDYSi4", overwrite = TRUE)
            template <- "KBASummary_Template_NewForm_NoQuestions_GeneralReview_National.docx"
          }
        }else if(reviewStage == "steering"){
          if(scope == "Global and National"){
            googledrive::drive_download("https://docs.google.com/document/d/1ztHExERMAN6GfgHeu1y2jwI7PPfuspjf", overwrite = TRUE)
            template <- "KBASummary_Template_NewForm_NoQuestions_SteeringCommittee_GlobalNational.docx"
          }else if(scope == "Global"){
            googledrive::drive_download("https://docs.google.com/document/d/1BIP6H5yJ9GZakuI9r2JmZq4Rzj0L-9Wz", overwrite = TRUE)
            template <- "KBASummary_Template_NewForm_NoQuestions_SteeringCommittee_Global.docx"
          }else if(scope == "National"){
            googledrive::drive_download("https://docs.google.com/document/d/1zzD8vb0X8kq2_B_lXwhoxqxcj8JK9IIe", overwrite = TRUE)
            template <- "KBASummary_Template_NewForm_NoQuestions_SteeringCommittee_National.docx"
          }
        }
      }else if(language == "french"){
        if(reviewStage == "technical"){
          if(scope == "Mondial et National"){
            googledrive::drive_download("https://docs.google.com/document/d/1NUno1-6dkFdMprf6RsmidllJxrKAyD1Z", overwrite = TRUE)
            template <- "KBASummary_Template_NewForm_NoQuestions_TechnicalReview_GlobalNational_FR.docx"
          }else if(scope == "Mondial"){
            googledrive::drive_download("https://docs.google.com/document/d/1NW2wqngvZvI-R3rr7oEeZKaVHY5ukU5m", overwrite = TRUE)
            template <- "KBASummary_Template_NewForm_NoQuestions_TechnicalReview_Global_FR.docx"
          }else if(scope == "National"){
            googledrive::drive_download("https://docs.google.com/document/d/1NISUHtepMym56gpsbUajqqBmTqhHZCrs", overwrite = TRUE)
            template <- "KBASummary_Template_NewForm_NoQuestions_TechnicalReview_National_FR.docx"
          }
        }else if(reviewStage == "general"){
          if(scope == "Mondial et National"){
            googledrive::drive_download("https://docs.google.com/document/d/1OQercOMhVcQiNsSXHwlRaytt81Z_PuCI", overwrite = TRUE)
            template <- "KBASummary_Template_NewForm_NoQuestions_GeneralReview_GlobalNational_FR.docx"
          }else if(scope == "Mondial"){
            googledrive::drive_download("https://docs.google.com/document/d/1NF4meIDvSh4r4GjU-laJkXpx4Wh0hDQc", overwrite = TRUE)
            template <- "KBASummary_Template_NewForm_NoQuestions_GeneralReview_Global_FR.docx"
          }else if(scope == "National"){
            googledrive::drive_download("https://docs.google.com/document/d/1OQTGhjCkivzS4xvUe07GrZockY5LtOJz", overwrite = TRUE)
            template <- "KBASummary_Template_NewForm_NoQuestions_GeneralReview_National_FR.docx"
          }
        }else if(reviewStage == "steering"){
          if(scope == "Mondial et National"){
            googledrive::drive_download("https://docs.google.com/document/d/1NXgSTVfKOIDT7WXMUATLSTZdAGz_YLGN", overwrite = TRUE)
            template <- "KBASummary_Template_NewForm_NoQuestions_SteeringCommittee_GlobalNational_FR.docx"
          }else if(scope == "Mondial"){
            googledrive::drive_download("https://docs.google.com/document/d/1O25cTlzetd5VCLW2wXDsdqN356XwwanN", overwrite = TRUE)
            template <- "KBASummary_Template_NewForm_NoQuestions_SteeringCommittee_Global_FR.docx"
          }else if(scope == "National"){
            googledrive::drive_download("https://docs.google.com/document/d/1NX2bERMqURwZNWmLdph6vrAI1e-TJAr_", overwrite = TRUE)
            template <- "KBASummary_Template_NewForm_NoQuestions_SteeringCommittee_National_FR.docx"
          }
        }
      }
    
      # Compute document name
      if(language == "english"){
        
        reviewStageLabel <- ifelse(reviewStage == "technical",
                                   "TechnicalReview",
                                   ifelse(reviewStage == "general",
                                          "GeneralReview",
                                          "SteeringCommittee"))
        
        doc <- paste0("Summary_", reviewStageLabel, "_", str_replace_all(string=nationalName, pattern=c(":| |\\(|\\)|/"), repl=""), "_", Sys.Date(), ".docx")
        
      }else{
        
        reviewStageLabel <- ifelse(reviewStage == "technical",
                                   "RévisionTechnique",
                                   ifelse(reviewStage == "general",
                                          "RévisionGénérale",
                                          "ComitéDeGestion"))
        
        doc <- paste0("Sommaire_", reviewStageLabel, "_", str_replace_all(string=nationalName, pattern=c(":| |\\(|\\)|/"), repl=""), "_", Sys.Date(), ".docx")
      }
      
      # Save
      doc <- renderInlineCode(template, doc)
      Sys.sleep(5)
      doc <- body_add_flextables(doc, doc, FT)
    
      KBAforms[step] <- doc
      
      # If the previous doc object was created, than the conversion was a success
      sucess <- TRUE
    
      # Compute the successful conversion into the table
      if(sucess){
        convert_res[step,"Result"] <- emo::ji("check")
        
        if(!formVersion %in% c(1, 1.1)){
          
          for(i in 1:nrow(species)){
            
            if(((!is.na(species$`Display taxon name?`[i]) && (species$`Display taxon name?`[i] == "No")) | (!is.na(species$`Display taxonomic group?`[i]) && (species$`Display taxonomic group?`[i] == "No")) | (!is.na(species$`Display assessment information?`[i]) && (species$`Display assessment information?`[i] == "No"))) & (!reviewStage == "general")){
              convert_res[step, "Message"] <- "WARNING: Contains unredacted sensitive information"
            }
          }
        }
      }
    }
  })
  
  # List to store the summaries AND the result table that will be displayed on the Shiny app 
  list_item <- list() # list to stock the summaries and a dataframe to see if it's a success or not
  list_item[[1]] <- KBAforms
  list_item[[2]] <- convert_res
  
  return(list_item)
}