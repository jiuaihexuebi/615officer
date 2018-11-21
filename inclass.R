library(officer)
library(magrittr)
library(tidyverse)
library(readxl)

#group member: Kecheng Liang, Bingqing Yu, Zhaobin Liu, Yudi Mao

##  pic a template you like
inclassexercise <- read_pptx("template.pptx") 

##  get the layout
layout_summary(inclassexercise)

master <- "Default Design"
layout_properties(x = inclassexercise, layout = "Title Slide", master = master )

#page 1
inclassexercise %>%  
  add_slide(layout = "Title Slide",master=master) %>% 
  ph_with_text(type="ctrTitle",str="Advantages of a Bear Market") %>% 
  ph_with_text(type="subTitle",str="Yes there is a positive side to a Bear Market!") 

#page 2
inclassexercise %>% 
  add_slide(layout = "Two Content", master = master) %>% 
  ph_with_text(type = "title", str = "Investing in Stocks ") %>%
  ph_with_ul(type = "body", index=1,
             str_list= c("Represents ownership in a firm","Earn a return in two ways",
                         "Price of the stock rises over time","Dividends are paid to the stockholder",
                         "Stockholders have claim on all assets"),
             level_list = c(1,1,2,2,1),
             style = fp_text(color = "black", font.size = 22)) %>% 
  ph_with_ul(type = "body", index=2,
             str_list= c("Right to vote for directors and on certain issues",
                         "Two types","Common stock","Right to vote","Receive dividends",
                         "Preferred stock","Receive a fixed dividend","Do not usually vote"),
             level_list = c(1,1,2,3,3,2,3,3),
             style = fp_text(color = "black", font.size = 22)) %>% 
  ph_with_text(type="ftr",str="11-2") %>% 
  ph_with_text(type="dt",str="Copyright 2006 Pearson Addison-Wesley. All rights reserved.")

#page 3
inclassexercise %<>% add_slide(layout = "Title and Content", master = master ) %>% 
  ph_with_text(type = "title", str = "Investing in Stocks: Sample Corporate Stock Certificate") %>% 
  ph_with_img( src = "page3.jpg")

#page 4
inclassexercise %<>% add_slide(layout = "Title and Content", master = master ) %>% 
  ph_with_text(type = "title", str = "What is a Bear Market?") %>% 
  ph_with_text(type = "body", str = "A decline of 15-20% of the broad market coupled with pessimistic sentiment underlying the stock market.")%>%
  ph_with_img_at( src = "page4.jpg", width = 5.5, height = 3.5, left = 3.5, top = 3.8)

#page 5
inclassexercise %<>% add_slide(layout = "Title and Content", master = master ) %>% 
  ph_with_text(type = "title", str = "Stock Market Indexes:the Dow Jones Industrial Average",index = 1) %>% 
  ph_with_img(src = "page5.png", height = 5,width = 9)

#page 6
inclassexercise %<>% add_slide(layout = "Title and Content", master = master ) %>% 
  ph_with_text(type = "title", str = "Dow Jones",index = 1) %>% 
  ph_with_img(src = "page6.png", height = 5,width = 9)

#page 7
inclassexercise %>%  add_slide(layout = "Title and Content", master = master) %>% 
  ph_with_text(type = "title", str = "The last Bear Market") %>% 
  ph_with_ul(
    type = "body", index = 1,
    str_list = c("Sept. 30, 2002  Dow  7528", 
                 "Jan. 5, 2004      Dow  10,568", 
                 "Oct. 8, 2007      Dow   14093"),
    level_list = c(1, 1, 1),
    style = fp_text(color = "black", font.size = 36) )

#page 8
inclassexercise %>% add_slide(layout = "Title and Content", master = master) %>% 
  ph_with_text(type = "title", str = "What do I do in a Bear Market") %>% 
  ph_with_ul(
    type = "body", index = 1,
    str_list = c("Decide whether this is a market correction or the start of something more", 
                 "Review the stocks you own", 
                 "Review stocks you wanted to own but were too expensive at time of research",
                 "Check your portfolio for balance or the type of stocks you own"),
    level_list = c(1, 1, 1, 1),
    style = fp_text(color = "black", font.size = 32) )

remove_slide(inclassexercise, index = 1)

print(inclassexercise, target = "exercise.pptx") 

