#library(reticulate)
library(xlsx)
#library(readxl)
library(tidyverse)
library(lubridate)
library(zoo)
library(scales)
library(reshape2)
library(treemapify)

# py_run_file("screenPy_tracker.py")

read.transposed.xlsx <- function(file, sheetName) {
  df <- read.xlsx(file, sheetName = sheetName , header = FALSE)
  dft <- as.data.frame(t(df[-1]), stringsAsFactors = FALSE) 
  names(dft) <- df[,1] 
  dft <- as.data.frame(lapply(dft,type.convert))
  return(dft)            
}

data1<-read.transposed.xlsx(paste("C:\\Users\\Aaron\\OneDrive - McGill University\\Programming\\screentime-tracker\\",
                           "StayFree Export - Total Usage - 8_20_21.xlsx", sep=""),
                           sheetName="Usage Time")
data1<-clean_data(data1, as.Date('08-01-20', "%m-%d-%y"), as.Date('08-20-21', "%m-%d-%y"))
data1<-data1[,1:10]
data2<-read.transposed.xlsx(paste("C:\\Users\\Aaron\\OneDrive - McGill University\\Programming\\screentime-tracker\\",
                                  "StayFree Export - Total Usage - 8_20_21 (2).xlsx", sep=""),
                            sheetName="Usage Time")
data2<-clean_data(data2, as.Date('08-01-20', "%m-%d-%y"), as.Date('08-20-21', "%m-%d-%y"))
data2<-data2[,1:12]

data3<-read.transposed.xlsx(paste("C:\\Users\\Aaron\\OneDrive - McGill University\\Programming\\screentime-tracker\\",
                                  "StayFree Export - Total Usage - 8_20_21 (early).xlsx", sep=""),
                            sheetName="Usage Time")
#data3<-clean_data(data3, as.Date('08-01-20', "%m-%d-%y"), as.Date('08-20-21', "%m-%d-%y"))
data3<-data3[,c(1:10,ncol(data3)-2)]

data4<-read.transposed.xlsx(paste("C:\\Users\\Aaron\\OneDrive - McGill University\\Programming\\screentime-tracker\\",
                                  "StayFree Export - Total Usage - 8_20_21 (late).xlsx", sep=""),
                            sheetName="Usage Time")
data4<-clean_data(data4, as.Date('08-01-20', "%m-%d-%y"), as.Date('08-20-21', "%m-%d-%y"))
data4<-data4[,1:12]

data5<-read.transposed.xlsx(paste("C:\\Users\\Aaron\\OneDrive - McGill University\\Programming\\screentime-tracker\\",
                                  "StayFree Export - Total Usage - 8_20_21 (late+1).xlsx", sep=""),
                            sheetName="Usage Time")
#data5<-clean_data(data5, as.Date('08-01-20', "%m-%d-%y"), as.Date('08-20-21', "%m-%d-%y"))
data5<-data5[,c(1:10,ncol(data5)-2)]

data6<-read.transposed.xlsx(paste("C:\\Users\\Aaron\\OneDrive - McGill University\\Programming\\screentime-tracker\\",
                                  "StayFree Export - Total Usage - 8_20_21 (later).xlsx", sep=""),
                            sheetName="Usage Time")
#data5<-clean_data(data5, as.Date('08-01-20', "%m-%d-%y"), as.Date('08-20-21', "%m-%d-%y"))
data6<-data6[,c(1:10,ncol(data6)-2)]

#view(data5)
bnd <-plyr::rbind.fill(data6, data3)
bnd[is.na(bnd)] <- "0s"
bnd <- unique(bnd)
bnd <- bnd[bnd$NA.!="Total Usage",]
bnd$Total.Usage2 <- sapply(bnd$Total.Usage, time_to_int)
bnd <- bnd[order(as.Date(bnd$NA., format="%B %d, %Y"),-bnd$Total.Usage2),]
bnd <- bnd[,!names(bnd)=="Total.Usage2"]
bnd <- bnd[!duplicated(bnd$NA.),]
view(bnd)

files <- c(paste("C:\\Users\\Aaron\\OneDrive - McGill University\\Programming\\screentime-tracker\\",
                 "StayFree Export - Total Usage - 8_1_21 (2).xlsx", sep=""),
           paste("C:\\Users\\Aaron\\OneDrive - McGill University\\Programming\\screentime-tracker\\",
                 "StayFree Export - Total Usage - 8_20_21 (early).xlsx", sep=""))
bnd <- data.frame()
for(i in 1:length(files)){
  bnd <- plyr::rbind.fill(bnd,
                          read.transposed.xlsx(
                            files[i], sheetName="Usage Time"))
  bnd[is.na(bnd)] <- "0s"
  bnd <- unique(bnd)
  bnd <- bnd[bnd$NA.!="Total Usage",]
  bnd$Total.Usage2 <- sapply(bnd$Total.Usage, time_to_int)
  bnd <- bnd[order(as.Date(bnd$NA., format="%B %d, %Y"),-bnd$Total.Usage2),]
  bnd <- bnd[,!names(bnd)=="Total.Usage2"]
  bnd <- bnd[!duplicated(bnd$NA.),]
}
view(bnd)

data<-read.files(files)
view(data)
datac <-clean_data(data, as.Date('07-13-20', "%m-%d-%y"), as.Date('08-10-21', "%m-%d-%y"),
                   tot=T)
warnings()
view(data[,-1])
view(datac)
read.files <- function(files){
  bnd <- data.frame()
  for(i in 1:length(files)){
    bnd <- plyr::rbind.fill(bnd,
                            read.transposed.xlsx(
                              files[i], sheetName="Usage Time"))
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
  if(tot){
    #data <- data[1:(nrow(data)-1), c(1,ncol(data)-2)]#####
    #data[,-1] <- sapply(data[,-1], time_to_int)
    data[,-1] <- apply(data[,-1], c(1,2), time_to_int)
    data$Total.Usage <- data$Total.Usage/3600
  } else{
    data <- data[,!names(data)=="Total.Usage"]
    # Manual edit because multiple apps have the name "Reminder"
    data <- rename(data, Reminder.2 = Reminder)
    data <- rename(data, Reminder = Reminder.1)
    data <- rename(data, Reminder.1 = Reminder.2)
    
    data[,-1] <- apply(data[,-1], c(1,2), time_to_int)
  }
  data[,1] <- as.Date(data[,1], "%B %d, %Y")
  data <- rename(data, Date = NA.)
  data <- rename_with(data, ~gsub(".", " ", .x, fixed=TRUE))
  data <- data[data$Date %in% start:end,]
  return(data)
}

clean_data_old <- function(data, start, end, tot=FALSE){
  # STEP 1 - get data into table
  # data <- read_xlsx(file_path, sheet = "Transposed", col_types = "numeric")
  # data <- rename(data, Reminder = Reminder...107)
  
  # OR
  
  #data <- read.xlsx(file_path, sheetName="Transposed", colClasses = "numeric")
  #data <- read.transposed.xlsx(file_path, sheetName="Usage Time")
  if(tot){
    data <- data[1:(nrow(data)-1), c(1,ncol(data)-2)]
    data[,-1] <- sapply(data[,-1], time_to_int)
    data$Total.Usage <- data$Total.Usage/3600
  } else{
    data <- data[1:(nrow(data)-1), 1:(ncol(data)-4)]
    data <- rename(data, Reminder.2 = Reminder)
    data <- rename(data, Reminder = Reminder.1)
    data[,-1] <- apply(data[,-1], c(1,2), time_to_int)
  }
  
  data[,1] <- as.Date(data[,1], "%B %d, %Y")
  data <- rename(data, Date = NA.)
  data <- rename_with(data, ~gsub(".", " ", .x, fixed=TRUE))

  # LAST DATE INCLUDED: 2020-05-05 (YYYY-MM-DD)
  #dates <- seq( start_date, end_date, by="days")
  #data <- data %>% add_column('Date'=dates, .before = 1)
  data <- data[data$Date %in% start:(end+1),]
  return(data)
}

prepare_data <- function(data, cutoff, tot=FALSE){
  data <- clean_data(data, tot)
  
  # SLICE ONLY DESIRED DATES NOW IF WANT A SMALLER PORTION
  # STEP 2 - melt
  data_melt <- data %>% melt(id = "Date", measure = c(-1)) %>% 
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

cutoff <- 10
data_melt <- prepare_data(data, cutoff)
ordered_apps_glob <- data_melt %>% group_by(App) %>% summarise(Avg = mean(Seconds)) %>% arrange(desc(Avg)) %>% pull(App)

# STEP 5 - plot
to_plot <- prep_to_plot(data_melt, ordered_apps_glob, "week", cutoff)
ggplot(to_plot, aes(Filtered_date, Daily_Avg_h, fill=Lumped_apps)) + 
  geom_area() + labs(x = "Month", y = "Daily time usage averaged over week (hours)") +
  scale_x_date(date_breaks = "1 month", labels = date_format("%m-%Y"))

data_tot <- clean_data(data, TRUE)
ggplot() +
  geom_bar(data=to_plot,
          aes(Filtered_date, Daily_Avg_h, fill=Lumped_apps),stat="identity") + 
  geom_line(data=data_tot, aes(Date, y=rollmean(`Total Usage`, 7, na.pad = T), colour='7 day rolling average')) +
  geom_line(data=data_tot, aes(Date, y=cummean(`Total Usage`), colour='Cumulative average')) +
  scale_colour_manual("", values = c("7 day rolling average"="black",
                                 "Cumulative average"="blue")) +
  labs(x = "Month", y = "Daily time usage averaged over week (hours)",
       title = "Daily phone time usage") +
  scale_x_date(date_breaks = "1 month", labels = date_format("%m-%Y"))

to_plot2 <- prep_to_plot_tree(data_melt, ordered_apps_glob, cutoff)
percent <- function(x, digits = 2, format = "f", ...) {
  paste0(formatC(100 * x, format = format, digits = digits, ...), "%")
}
ggplot(to_plot2, aes(area=sum, fill=Lumped_apps, label=percent(sum/sum(sum)))) +
  geom_treemap() + geom_treemap_text()


to_plot3 <- prep_to_plot_week(data_melt, ordered_apps_glob, cutoff)
ggplot(to_plot3, aes(x=Weekday, y=Daily_Avg_h, fill=Lumped_apps)) +
  geom_bar(stat="identity")
  