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

data<-read.transposed.xlsx(paste(
                           "StayFree Export - Total Usage - 6_25_21.xlsx", sep=""),
                           sheetName="Usage Time")

clean_data <- function(data, tot=FALSE){
  # STEP 1 - get data into table
  # data <- read_xlsx(file_path, sheet = "Transposed", col_types = "numeric")
  # data <- rename(data, Reminder = Reminder...107)
  
  # OR
  
  #data <- read.xlsx(file_path, sheetName="Transposed", colClasses = "numeric")
  #data <- read.transposed.xlsx(file_path, sheetName="Usage Time")
  if(tot){
    data <- data[1:(nrow(data)-2), c(1,ncol(data)-3)]
    data$Total.Usage <- data$Total.Usage/3600
  } else{
    data <- data[1:(nrow(data)-2), 1:(ncol(data)-4)]
    data <- rename(data, Reminder.2 = Reminder)
    data <- rename(data, Reminder = Reminder.1)
  }
  
  data[,1] <- as.Date(data[,1], "%B %d, %Y")
  data <- rename(data, Date = NA.)
  data <- rename_with(data, ~gsub(".", " ", .x, fixed=TRUE))

  # LAST DATE INCLUDED: 2020-05-05 (YYYY-MM-DD)
  #dates <- seq( start_date, end_date, by="days")
  #data <- data %>% add_column('Date'=dates, .before = 1)
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

?geom_treemap_text
to_plot3 <- prep_to_plot_week(data_melt, ordered_apps_glob, cutoff)
ggplot(to_plot3, aes(x=Weekday, y=Daily_Avg_h, fill=Lumped_apps)) +
  geom_bar(stat="identity")