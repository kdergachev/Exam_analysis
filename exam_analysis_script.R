rm(list = ls())
setwd("D:/RScripts/exam_analysis")

#                  !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
#                         !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
#                                           !!!!!!!!!!!!!!!!!!!!!!!!!!!
# !!!!!!!!!!!!Does not work as a whole. Get tempvariable first (line43), then continue!!!!!!!!!!!!!!!
#                                           !!!!!!!!!!!!!!!!!!!!!!!!!!!    
#                         !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
#                  !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

open_exam_results_xl <- function(fpath){
  
  #extract results from xlsx file
  
  #get the table itself
  
  library(readxl)
  results <- read_excel(path = fpath, na = c("", "abs", "н/я"))
  View(results)
  
  #get coordinate data to extract results
  
  coord <- (readline("Input (separated by spaces) namecolumn, resultcolumn, upper and lower bound rows"))
  coord <- as.integer(unlist(strsplit(coord, " ")))
  name <- readline("input the name in correct format or result")
  if (is.na(as.numeric(name))){
    print('rrr')
    MyRes <- as.numeric(results[which(results[coord[1]] == name), coord[2]])
  }else if (!is.na(as.numeric(name))){
    print('ttt')
    MyRes <- as.numeric(name)
  }
  
  
  #extract results vector
  
  results <- as.numeric(unlist(results[coord[3]:coord[4], coord[2]]))
  ret <- list("res" = MyRes, "vec" = results)
  return(ret)
}

tempvariable <- open_exam_results_xl("Project and HA grades (1).xlsx")


tempvariable$vec_no_na <- na.omit(tempvariable$vec)
tempvariable$vec[is.na(tempvariable$vec)] <- 0
names(tempvariable)[2] <- "vec_na_is0"

# now tempvar == Myresult and 2 resultarrays where one is with NAs removed and another one has NA = 0

result_dt <- function(mres, dtaNArem, dtaNA0){
  
  # turn tempvar contents into a summary
  
  # df template to be filled
  out <- data.frame(matrix(ncol = 2, nrow = 10))
  rownames(out) <- c("myres", "min", "1st Q.", "median", "3rd Q.", "max", "mean", "size", "mypos", "myquant")
  colnames(out) <- c("NA_removed", "NA_are_zeros")
  
  #filling it in
  out["myres", ] <- c(mres, mres)
  j <- 1 # 1 fills no NA column, 2 fills NA = 0 one
  for (i in list(dtaNArem, dtaNA0)){
    # the easy ones
    out["size", j] <- length(i)
    out["mypos", j] <- paste(sum(i < mres) + 1, sum(i <= mres), sep = " - ")
    out["myquant", j] <- sum(i < mres)/as.integer(out["size", j])
    tt <- c(1, 2, 3, 5, 6, 4) # positions 1, 2, 3, 5, 6, 4 correspond to 2:7 (1:6 + 1) of the df
    for (t in c(1:6)){
      # extracting summary()
      out[t+1, j] <- summary(i)[tt[t]] # overcomplicated but works
    }
    j <- j + 1
  }
  return(out)
}

output <- result_dt(tempvariable$res, tempvariable$vec_no_na, tempvariable$vec_na_is0)

print(output)

# plot density
plot(density(tempvariable$vec_na_is0), xlab = "exam mark")
abline(v = tempvariable$res)
axis(side = 1, at = seq(-20, 100, by = 10))

plot(density(tempvariable$vec_no_na), xlab = "exam mark")
abline(v = tempvariable$res)
axis(side = 1, at = seq(-20, 100, by = 10))












###{
rm(list = ls())
setwd("D:/RScripts/exam_analysis")

inputs <- list("filename" = "ICEF-Winter-2019-2020.xlsx", 
               "pos" = c(2, 0, 13, 2), "MyRes" = 50, 
               "name" = "Дергачев Кирилл Олегович") #example and initial input


show_exam_results <- function(filename, startrow, indexcol, colres, 
                              colnms, Name = "Дергачев Кирилл Олегович"){
  #change inputs
  #fname, (first numeric row, index to calc total number, result column, name column), name
  #if pos[2] = 0 -> count names, if anything else -> max(leftmost column(ordinal one))
  
  inputs$filename <- filename
  inputs$pos[3] <- colres
  inputs$pos[4] <- colnms
  inputs$pos[1] <- startrow
  inputs$pos[2] <- indexcol
  inputs$name <- Name
  output <- list(exam = strsplit(inputs$filename, ".", fixed = TRUE)[[1]][1])
  #parse xlsx file
  if (grepl(".xlsx", inputs$filename) == TRUE){
    library("readxl")
    results <- read_excel(path = inputs$filename, na = c(""))
    if (inputs$pos[2] != 0){
      inputs$pos[2] <- max((unlist(results[1])), na.rm = TRUE)
    }else{
      inputs$pos[2] <- length(unlist(results[inputs$pos[4]]))
    }
    inputs$MyRes <- as.numeric(results[which(results[[inputs$pos[4]]] == inputs$name), inputs$pos[3]])
    results <- results[inputs$pos[1]:inputs$pos[2] ,inputs$pos[3]]
    results[results == "н/я"] <- NA
    results[results == "abs"] <- NA
    #tibble of all final results recieved, now to analyse:
    
    
    #na dropped
    lastrow1 <- inputs$pos[2] - inputs$pos[1] + 1 - sum(is.na(results))
    results_no_na <- na.omit(results)
    results_no_na <- c(matrix(unlist(results_no_na[1]), ncol = 1, nrow = lastrow1))
    results_no_na <- as.integer(results_no_na)
    
    
    #na == 0
    lastrow2 <- inputs$pos[2] - inputs$pos[1] + 1
    results_na <- rapply(results, f = function(x) ifelse(is.na(x),0,x), how = "replace")
    results_na <- c(matrix(unlist(results_na[1]), ncol = 1, nrow = lastrow2))
    results_na <- as.integer(results_na)
    output$MyRes <- inputs$MyRes
    
    
    output$summary_no_na <- summary(results_no_na)
    output$my_pos_and_overall_num_no_na <- c(sum(results_no_na < inputs$MyRes), lastrow1) 
    output$my_quant_no_na <- sum(results_no_na < inputs$MyRes)/lastrow1
    output$plot_density_no_na <- plot(density(results_no_na))
    
    output$summary_na <- summary(results_na)
    output$my_pos_and_overall_num_na <- c(sum(results_na < inputs$MyRes), lastrow2)
    output$my_quant_na <- sum(results_na < inputs$MyRes)/lastrow2
    output$plot_density_na <- plot(density(results_na))
    #output created for 1st case, now to the second (results are in pdf)
    
  }else if (grepl(".pdf", inputs$filename) == TRUE){
    
    
    library(tabulizer)
    library(dplyr)
    
    results <- extract_tables("Banking.pdf")
    res <- matrix(nrow = 0, ncol = dim(results[[1]])[2])
    
    for (i in 1:length(results)){
      results[[i]] <- matrix(unlist(results[[i]]), nrow = length(results[[i]][, 1]), byrow = F)
      res <- rbind(res, results[[i]])
    }
    #inputs$MyRes <- as.numeric(res[which(res[inputs$pos[4]] == inputs$name), inputs$pos[3]])
    inputs$MyRes <- 15.5
    res <- as.numeric(sub(",", ".", res[2:length(res[, 1]), inputs$pos[3]], fixed = TRUE))
    res_na <- res
    res_na[is.na(res_na) == TRUE] <- 0
    res_no_na <- na.omit(res)
    
    
    output$summary_no_na <- summary(res_no_na)
    output$my_pos_and_overall_num_no_na <- c(sum(res_no_na<inputs$MyRes), length(res_no_na))
    output$my_quant_no_na <- sum(res_no_na<inputs$MyRes)/length(res_no_na)
    output$plot_density_no_na <- plot(density(res_no_na))
    
    output$summary_na <- summary(res_na)
    output$my_pos_and_overall_num_na <- c(sum(res_na<inputs$MyRes), length(res_na))
    output$my_quant_na <- sum(res_na<inputs$MyRes)/length(res_na)
    output$plot_density_na <- plot(density(res_na))
  }
  names(output) <- paste(output$exam, names(output), sep = "_")
  return(output)
  rm(list == ls())
}


show_exam_results("ICEF-Winter-2019-2020.xlsx", 2, 0, 32, 2)
}