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

clean_data <- function(data, tot=FALSE){
  if(tot){
    data <- data[1:(nrow(data)-2), c(1,ncol(data)-2)]
    data[,-1] <- sapply(data[,-1], time_to_int)
    data$Total.Usage <- data$Total.Usage/3600
  } else{
    data <- data[1:(nrow(data)-2), 1:(ncol(data)-4)]
    # Manual edit because multiple apps have the name "Reminder"
    data <- rename(data, Reminder.2 = Reminder)
    data <- rename(data, Reminder = Reminder.1)
    data[,-1] <- apply(data[,-1], c(1,2), time_to_int)
  }
  data[,1] <- as.Date(data[,1], "%B %d, %Y")
  data <- rename(data, Date = NA.)
  data <- rename_with(data, ~gsub(".", " ", .x, fixed=TRUE))
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

percent <- function(x, digits = 2, format = "f", ...) {
  paste0(formatC(100 * x, format = format, digits = digits, ...), "%")
}





function(input, output) {
  data<-eventReactive(input$file, {
    read.transposed.xlsx(input$file$datapath, sheetName="Usage Time")
  })

  data_melt <- eventReactive({
    data()
    input$numApps}, {
    prepare_data(data(), input$numApps)
  })
  ordered_apps_glob <- eventReactive(data_melt(), {
    data_melt() %>% group_by(App) %>% summarise(Avg = mean(Seconds)) %>%
      arrange(desc(Avg)) %>% pull(App)
  })
  
  # 1 - AREA GRAPH
  to_plot <- eventReactive(ordered_apps_glob(), {
    prep_to_plot(data_melt(), ordered_apps_glob(), "week", input$numApps)
  })#, ignoreInit = TRUE)
  # 2 - TREEMAP
  to_plot2 <- eventReactive(ordered_apps_glob(), {
    prep_to_plot_tree(data_melt(), ordered_apps_glob(), input$numApps)
  })
  # 3 - BAR GRAPH BY WEEKDAY
  to_plot3 <- eventReactive(ordered_apps_glob(), {
    prep_to_plot_week(data_melt(), ordered_apps_glob(), input$numApps)
  })
  # 4 - WEEKLY BAR GRAPH OVER TIME
  data_tot <- eventReactive(data(), {
    clean_data(data(), TRUE)
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
  
  # output$plot1 <- renderPlot({
  #   par(mar = c(5.1, 4.1, 0, 1))
  #   ggplot(to_plot(), aes(Filtered_date, Daily_Avg_h, fill=Lumped_apps)) + 
  #     geom_area() + labs(x = "Month", y = "Daily time usage averaged over week (hours)") +
  #     scale_x_date(date_breaks = "1 month", labels = date_format("%m-%Y"))
  # })
  output$plot2 <- renderPlot({
    ggplot(to_plot2(), aes(area=sum, fill=Lumped_apps, label=percent(sum/sum(sum)))) +
      geom_treemap() + geom_treemap_text()
  })
  output$plot3 <- renderPlot({
    ggplot(to_plot3(), aes(x=Weekday, y=Daily_Avg_h, fill=Lumped_apps)) +
      geom_bar(stat="identity")
  })
  output$plot4 <- renderPlot({
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
  })
  output$pca <- renderPlot({
    if(input$numApps<=1){return ()}
    ggplot_pca(pca(), labels = c("Sunday", "Monday", "Tuesday", "Wednesday",
                               "Thursday", "Friday", "Saturday"))
  })
  output$contrib <- renderPlot({
    if(input$numApps<=2){return ()}
    ggplot(contrib(), aes(x=stat, fill=variable, y=value)) +
      geom_bar(stat="identity", position=PositionDodge) +
      facet_grid(~variable) + theme(legend.position = "top",
                                    axis.text.x = element_text(angle=90))+coord_flip()
  })
}
