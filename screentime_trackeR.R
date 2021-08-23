library(readxl)
#library(xlsx)
library(tidyverse)
library(lubridate)
library(zoo)
library(scales)
library(reshape2)
library(treemapify)
library(GGally)
library(AMR)
#library(FactoMineR)

read.files <- function(files){
  bnd <- data.frame()
  for(i in 1:length(files[,1])){
    bnd <- plyr::rbind.fill(bnd,
                            read.transposed.xlsx(
                              files[[i, 'datapath']], sheetName="Usage Time"))
    bnd[is.na(bnd)] <- "0s"
    bnd <- unique(bnd)
    bnd <- bnd[bnd$NA.!="Total Usage",]
    bnd$Total.Usage2 <- sapply(bnd$Total.Usage, time_to_int)
    bnd <- bnd[order(as.Date(bnd$NA., format="%B %d, %Y"),-bnd$Total.Usage2),]
    bnd <- bnd[!duplicated(bnd$NA.),]
    bnd <- bnd[,!names(bnd)=="Total.Usage2" &
                 !names(bnd)=="Created.by..StayFree.."]
    bnd <- bnd %>% select(!starts_with("Creation.date.."))
  }
  bnd
}

read.transposed.xlsx <- function(file, sheetName) {
  df <- read.xlsx(file, sheetName = sheetName , header = FALSE)
  dft <- as.data.frame(t(df[-1]), stringsAsFactors = FALSE) 
  names(dft) <- df[,1] 
  dft <- as.data.frame(lapply(dft,type.convert))
  return(dft)            
}

time_to_int <- function(time){
  time <- as.character(time)
  h<-0
  m<-0
  s<-0
  if(grepl("h", time, fixed=T)){
    h <- str_extract(time, "\\d{1,2}h")
    h <- as.integer(substr(h, 1, nchar(h)-1))
  }
  if(grepl("m", time, fixed=T)){
    m <- str_extract(time, "\\d{1,2}m")
    m <- as.integer(substr(m, 1, nchar(m)-1))
  }
  if(grepl("s", time, fixed=T)){
    s <- str_extract(time, "\\d{1,2}s")
    s <- as.integer(substr(s, 1, nchar(s)-1))
  }
  return(h*3600 + m*60 + s) 
}

clean_data <- function(data, start, end, tot=FALSE){
  data[,-1] <- apply(data[,-1], c(1,2), time_to_int)
  if(tot){
    data$Total.Usage <- data$Total.Usage/3600
  } else{
    data <- data[,!names(data)=="Total.Usage"]
    # Manual edit because multiple apps have the name "Reminder"
    data <- rename(data, Reminder.2 = Reminder)
    data <- rename(data, Reminder = Reminder.1)
    data <- rename(data, Reminder.1 = Reminder.2)
  }
  data[,1] <- as.Date(data[,1], "%B %d, %Y")
  data <- rename(data, Date = NA.)
  data <- rename_with(data, ~gsub(".", " ", .x, fixed=TRUE))
  data <- data[data$Date %in% start:end,]
  return(data)
}

prepare_data <- function(data, cutoff, start, end, tot=FALSE){
  # STEP 1 - clean
  clean_data <- clean_data(data, start, end, tot)
  # STEP 2 - melt
  data_melt <- clean_data %>% melt(id = "Date", measure = c(-1)) %>% 
    rename(App=variable, Seconds=value) %>% mutate(App=as.character.factor(App))
  # Add a weekday column
  data_melt$Weekday <- weekdays(data_melt$Date)
  # STEP 3 - lump less used apps
  ordered_apps <- data_melt %>% group_by(App) %>%
    summarise(Avg = mean(Seconds)) %>% arrange(desc(Avg)) %>% pull(App)
  print(ordered_apps[1:cutoff])
  data_melt <- data_melt %>% group_by(App) %>%
    mutate(Lumped_apps =
             ifelse(App %in% ordered_apps[1:cutoff], App, "Other"))
  return(data_melt)
}

prep_to_plot <- function(data_melt, ordered_apps, filter_by, cutoff){
  # STEP 4 - add date filters
  data_sum <- data_melt %>% group_by(Filtered_date=floor_date(Date, filter_by), Lumped_apps) %>%
    summarise(sum = sum(Seconds))
  n_col <- data_melt %>% group_by(Filtered_date=floor_date(Date, filter_by), Lumped_apps) %>% 
    count() %>% 
    mutate(n=ifelse(Lumped_apps == "Other", n/(length(ordered_apps)-cutoff), n)) %>%
    pull(n)
  data_sum <- data_sum %>% add_column(n=n_col)
  to_plot <- data_sum %>% mutate(Daily_Avg = sum/n) %>%
    select(Filtered_date, Lumped_apps, Daily_Avg)
  # Convert to factors
  to_plot$Lumped_apps <- factor(to_plot$Lumped_apps, c(rev(ordered_apps[1:cutoff]),"Other"))
  # Add an hours column
  to_plot$Daily_Avg_h <- to_plot$Daily_Avg/3600
  return(to_plot)
}

prep_to_plot_tree <- function(data_melt, ordered_apps, cutoff){
  # STEP 4 - group together by sum
  to_plot <- data_melt %>% group_by(Lumped_apps) %>%
    summarise(sum = sum(Seconds))
  to_plot$Lumped_apps <- factor(to_plot$Lumped_apps, c(rev(ordered_apps[1:cutoff]),"Other"))
  return(to_plot)
}

prep_to_plot_week <- function(data_melt, ordered_apps, cutoff){
  # STEP 4 - add weekday filter
  data_sum <- data_melt %>% group_by(Weekday, Lumped_apps) %>%
    summarise(sum = sum(Seconds))
  n_col <- data_melt %>% group_by(Weekday, Lumped_apps) %>% 
    count() %>% 
    mutate(n=ifelse(Lumped_apps == "Other", n/(length(ordered_apps)-cutoff), n)) %>%
    pull(n)
  data_sum <- data_sum %>% add_column(n=n_col)
  to_plot <- data_sum %>% mutate(Daily_Avg = sum/n) %>%
    select(Weekday, Lumped_apps, Daily_Avg)
  to_plot$Lumped_apps <- factor(to_plot$Lumped_apps, c(rev(ordered_apps[1:cutoff]),"Other"))
  to_plot$Weekday <- factor(to_plot$Weekday,
                            c("Sunday", "Monday", "Tuesday", "Wednesday",
                              "Thursday", "Friday", "Saturday"))
  # Add an hours column
  to_plot$Daily_Avg_h <- to_plot$Daily_Avg/3600
  return(to_plot)
}

percent <- function(x, digits = 2, format = "f", ...) {
  paste0(formatC(100 * x, format = format, digits = digits, ...), "%")
}

dateDiff <- 21