if(formVersion %in% c(1, 1.1)){
  check_checkboxes %<>% .[c(1:5,7:nrow(.)),] # Cell N8 is obsolete in v1.1 of the Proposal Form (it doens't link to any actual checkbox)
}else{
  print("??")
}