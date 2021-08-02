library(xlsx)
library(tidyverse)
library(lubridate)
library(zoo)
library(scales)
library(reshape2)
library(treemapify)
library(GGally)
library(AMR)
#library(FactoMineR)
library(shiny)

read.transposed.xlsx <- function(file, sheetName) {
  df <- read.xlsx(file, sheetName = sheetName , header = FALSE)
  dft <- as.data.frame(t(df[-1]), stringsAsFactors = FALSE) 
  names(dft) <- df[,1] 
  dft <- as.data.frame(lapply(dft,type.convert))
  return(dft)            
}

time_to_int <- function(time){
  time <- as.integer(rev(strsplit(as.character(time), "[hms]\\s?")[[1]]))
  time3 <- c(rep(0,3))
  time3 <- c(time, time3)
  return(time3[1]+time3[2]*60+time3[3]*3600)
}

clean_data <- function(data, start=NULL, end=NULL, tot=FALSE){
  if(tot){
    data <- data[1:(nrow(data)-1), c(1,ncol(data)-2)]
    data[,-1] <- sapply(data[,-1], time_to_int)
    data$Total.Usage <- data$Total.Usage/3600
  } else{
    data <- data[1:(nrow(data)-1), 1:(ncol(data)-4)]
    # Manual edit because multiple apps have the name "Reminder"
    data <- rename(data, Reminder.2 = Reminder)
    data <- rename(data, Reminder = Reminder.1)
    data <- rename(data, Reminder.1 = Reminder.2)
    data[,-1] <- apply(data[,-1], c(1,2), time_to_int)
  }
  data[,1] <- as.Date(data[,1], "%B %d, %Y")
  data <- rename(data, Date = NA.)
  data <- rename_with(data, ~gsub(".", " ", .x, fixed=TRUE))
  # if(start!=end){
  #   data <- data[data$Date %in% start:end,]
  # }
  data <- data[data$Date %in% start:(end+1),]
  return(data)
}

prepare_data <- function(data, cutoff, start=NULL, end=NULL, tot=FALSE){
  # STEP 1 - clean
  data <- clean_data(data, start, end, tot)
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

percent <- function(x, digits = 2, format = "f", ...) {
  paste0(formatC(100 * x, format = format, digits = digits, ...), "%")
}

dateDiff <- 21


function(input, output) {
  data<-eventReactive(input$file, {
    read.transposed.xlsx(input$file$datapath, sheetName="Usage Time")
  })
  data_melt <- eventReactive({
    data()
    input$numApps
    input$startDate
    input$endDate}, {
    prepare_data(data(), input$numApps, input$startDate, input$endDate)
  })
  ordered_apps_glob <- eventReactive(data_melt(), {
    data_melt() %>% group_by(App) %>% summarise(Avg = mean(Seconds)) %>%
      arrange(desc(Avg)) %>% pull(App)
  })
  startDate <- eventReactive(data(), {
    as.Date(data()[1,1], "%B %d, %Y")
  })#, ignoreInit=T)
  endDate <- eventReactive(data(), {
    as.Date(data()[nrow(data())-1,1], "%B %d, %Y")
  })#, ignoreInit=T)
  
  # 1 - AREA GRAPH
  to_plot <- eventReactive(ordered_apps_glob(), {
    if(input$endDate-input$startDate>dateDiff){#} || input$endDate==input$startDate){
      prep_to_plot(data_melt(), ordered_apps_glob(), "week", input$numApps)
    } else {
      prep_to_plot(data_melt(), ordered_apps_glob(), "day", input$numApps)
    }
  })
  # 2 - TREEMAP
  to_plot2 <- eventReactive(ordered_apps_glob(), {
    prep_to_plot_tree(data_melt(), ordered_apps_glob(), input$numApps)
  })
  # 3 - BAR GRAPH BY WEEKDAY
  to_plot3 <- eventReactive(ordered_apps_glob(), {
    prep_to_plot_week(data_melt(), ordered_apps_glob(), input$numApps)
  })
  # 4 - WEEKLY BAR GRAPH OVER TIME
  data_tot <- eventReactive({
    data()
    input$startDate
    input$endDate}, {
    clean_data(data(), input$startDate, input$endDate, tot=TRUE)
  })
  # 5 - PCA
  pca <- eventReactive(to_plot3(), {
    weekday_data <- dcast(to_plot3(), Weekday~Lumped_apps, value.var = "Daily_Avg_h")
    melted_weekday <- to_plot3()[c(1:2,4)]
    prcomp(weekday_data[2:(input$numApps+1)])
  })
  # 6 - CONTRIB
  contrib <- eventReactive(pca(), {
    contrib <- data.frame(stat=rownames(pca()$rotation),pca()$rotation[,1:3])
    contrib$stat <- factor(contrib$stat, levels = contrib$stat[order(contrib$PC2^2)])
    melt(contrib)
  })
  
  output$numApps <- renderUI({
    if(is.null(input$file)){return()}
    sliderInput('numApps', 'Number of apps', value = min(10, ncol(data())/2),
                min = 2, max = min(20, ncol(data())))
  })
  output$startDate <- renderUI({
    dateInput('startDate', 'Start date', value=startDate(),
              min=startDate(), max=input$endDate)
  })
  output$endDate <- renderUI({
    dateInput('endDate', 'End date', value=endDate(),
              min=input$startDate, max=endDate())
  })
  
  # output$plot1 <- renderPlot({
  #   par(mar = c(5.1, 4.1, 0, 1))
  #   ggplot(to_plot(), aes(Filtered_date, Daily_Avg_h, fill=Lumped_apps)) + 
  #     geom_area() + labs(x = "Month", y = "Daily time usage averaged over week (hours)") +
  #     scale_x_date(date_breaks = "1 month", labels = date_format("%m-%Y"))
  # })
  output$plot2 <- renderPlot({
    if(is.null(input$file)){return()}
    ggplot(to_plot2(), aes(area=sum, fill=Lumped_apps, label=percent(sum/sum(sum)))) +
      geom_treemap() + geom_treemap_text()
  })
  output$plot3 <- renderPlot({
    if(is.null(input$file)){return()}
    ggplot(to_plot3(), aes(x=Weekday, y=Daily_Avg_h, fill=Lumped_apps)) +
      geom_bar(stat="identity")
  })
  output$plot4 <- renderPlot({
    if(is.null(input$file)){return()}
    if(input$endDate-input$startDate>dateDiff){#} || input$endDate==input$startDate){
      ggplot() +
        geom_bar(data=to_plot(),
                 aes(Filtered_date, Daily_Avg_h, fill=Lumped_apps),stat="identity") +
        geom_line(data=data_tot(), aes(Date, y=rollmean(`Total Usage`, 7, na.pad = T), colour='7 day rolling average')) +
        geom_line(data=data_tot(), aes(Date, y=cummean(`Total Usage`), colour='Cumulative average')) +
        scale_colour_manual("", values = c("7 day rolling average"="black",
                                           "Cumulative average"="blue")) +
        labs(x = "Month", y = "Daily time usage averaged over week (hours)",
             title = "Daily phone time usage") +
        scale_x_date(date_breaks = "1 month", labels = date_format("%m-%y")) 
    } else {
      ggplot() +
        geom_bar(data=to_plot(),
                 aes(Filtered_date, Daily_Avg_h, fill=Lumped_apps),stat="identity") +
        geom_line(data=data_tot(), aes(Date, y=cummean(`Total Usage`), colour='Cumulative average')) +
        scale_colour_manual("", values = c("Cumulative average"="blue")) +
        labs(x = "Day", y = "Daily time usage",
             title = "Daily phone time usage") +
        scale_x_date(date_breaks = "1 day", labels = date_format("%m-%d"))
    }
  })
  output$pca <- renderPlot({
    if(is.null(input$file) || input$numApps<=1 || input$endDate-input$startDate<7){return ()}
    ggplot_pca(pca(), labels = c("Sunday", "Monday", "Tuesday", "Wednesday",
                               "Thursday", "Friday", "Saturday"))
  })
  output$contrib <- renderPlot({
    if(is.null(input$file) || input$numApps<=2 || input$endDate-input$startDate<7){return ()}
    ggplot(contrib(), aes(x=stat, fill=variable, y=value)) +
      geom_bar(stat="identity", position=PositionDodge) +
      facet_grid(~variable) + theme(legend.position = "top",
                                    axis.text.x = element_text(angle=90))+coord_flip()
  })
}
