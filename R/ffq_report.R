# 这是膳食反馈报告
#

ffq_report <- function(data = data, dir= dir) {

  options(warn = -1)

  if (suppressPackageStartupMessages(!require(tidyverse)))  {
    install.packages('tidyverse')
    suppressPackageStartupMessages(library(tidyverse))
  }

  if (suppressPackageStartupMessages(!require(formattable)))  {
    install.packages('formattable')
    suppressPackageStartupMessages(library(formattable))
  }

  if (suppressPackageStartupMessages(!require(htmltools)))  {
    install.packages('htmltools')
    suppressPackageStartupMessages(library(htmltools))
  }

  if (suppressPackageStartupMessages(!require(haven)))  {
    install.packages('haven')
    suppressPackageStartupMessages(library(haven))
  }

  if (suppressPackageStartupMessages(!require(webshot)))  {
    install.packages('webshot')
    suppressPackageStartupMessages(library(webshot))
  }

  if (suppressPackageStartupMessages(!require(htmltools)))  {
    install.packages('htmltools')
    suppressPackageStartupMessages(library(htmltools))
  }

  if (suppressPackageStartupMessages(!require(officer)))  {
    install.packages('officer')
    suppressPackageStartupMessages(library(officer))
  }

  if (suppressPackageStartupMessages(!require(doconv)))  {
    install.packages('doconv')
    suppressPackageStartupMessages(library(doconv))
  }

  data <- read_dta(file = data)

  diet_data <- data %>% select(-id, -name)

  frequency <- diet_data %>% select(ends_with('b1'))

  for (i in 1:nrow(frequency)) {
    for (j in 1:ncol(frequency)) {
    if (frequency[i,j]==1) {frequency[i,j]=frequency[i,j]/365} else if (frequency[i,j]==frequency[i,j]) {frequency[i,j]=frequency[i,j]/30} else if (frequency[i,j]==3) {frequency[i,j]=frequency[i,j]/7} else {frequency[i,j]=frequency[i,j]}
    }
  }

  quantify <- diet_data %>% select(ends_with('b2')) %>% mutate('b79b2'=1) %>% relocate(b79b2,.after = b78b2)

  daily_quantify <- as.matrix(frequency) * as.matrix(quantify)

  index <- c(1,1,1,1,1,1,1,1,40,1,
             1,1,1,1,1,1,1,1,1,1,
             1,1,1,1,1,1,1,1,1,1,
             1,1,8,1,1,1,1,1,1,1,
             1,1,1,1,1,1,1,1,1,1,
             1,1,18,1,1,1,1,1,1,1,
             1,250,250,8,8,150,150,21,200,1,
             1,1,1,1,1,1,1,1,16,50,
             50,50,50
             )

  daily_quantify <- t(daily_quantify)*index

  daily_quantify <- data.frame(daily_quantify)

  food <- daily_quantify[79,] %>% apply(2,sum) %>% as.data.frame() %>% t() %>% as.data.frame() %>% mutate(group='油 (g/d)') %>% bind_rows(
         daily_quantify[62:66,] %>% apply(2,sum) %>% as.data.frame() %>% t() %>% as.data.frame() %>% mutate(group='奶类 (g/d)'),
         daily_quantify[c(18:19,71),] %>% apply(2,sum) %>% as.data.frame() %>% t() %>% as.data.frame()%>% mutate(group='大豆/坚果 (g/d)'),
         daily_quantify[c(44:61),] %>% apply(2,sum) %>% as.data.frame() %>% t() %>% as.data.frame() %>% mutate(group='肉禽蛋鱼类 (g/d)'),
         daily_quantify[c(44:53),] %>% apply(2,sum) %>% as.data.frame() %>% t() %>% as.data.frame() %>% mutate(group='禽肉类 (g/d)'),
         daily_quantify[c(54:60),] %>% apply(2,sum) %>% as.data.frame() %>% t() %>% as.data.frame() %>% mutate(group='鱼虾类 (g/d)'),
         daily_quantify[c(61),] %>% apply(2,sum) %>% as.data.frame() %>% t() %>% as.data.frame() %>% mutate(group='蛋类 (g/d)'),
         daily_quantify[c(21:33),] %>% apply(2,sum) %>% as.data.frame() %>% t() %>% as.data.frame() %>% mutate(group='蔬菜类 (g/d)'),
         daily_quantify[c(34:43),] %>% apply(2,sum) %>% as.data.frame() %>% t() %>% as.data.frame() %>% mutate(group='水果类 (g/d)'),
         daily_quantify[c(1:3),] %>% apply(2,sum) %>% as.data.frame() %>% t() %>% as.data.frame() %>% mutate(group='谷类 (g/d)'),
         daily_quantify[78,] %>% apply(2,sum) %>% as.data.frame() %>% t() %>% as.data.frame() %>% mutate(group='水 (ml/d)')
         )

  names(food)[1:length(data$id)] <- data$id
  names(daily_quantify)[1:length(data$id)] <- data$id

  component <- read_excel("食物成分.xls",
                           skip = 1)

  for (i in 1:length(data$id)) {
    temp <- food %>% select(group, data[i,]$id) %>% mutate(备孕及孕早期孕妇推荐摄入量=c('<25g', '300g', '25g', '130-180g', '40-65g', '40-65g', '40-50g', '300-500g', '200-300g', '200-250g','1500-1700ml'))

    names(temp)[2] <- '摄入量'

    food_table <- temp %>% filter(group=='油 (g/d)') %>% mutate(min=c(25), 说明=case_when(摄入量>min~'摄入偏高，推荐适当降低摄入', TRUE~'摄入合理')) %>% bind_rows(
      temp %>% filter(group=='蛋类 (g/d)') %>% mutate(min=40,max=50, 说明=case_when(摄入量>max~'摄入偏高，推荐适当降低摄入', 摄入量<min~'摄入偏低，推荐适当增加摄入', TRUE~'摄入合理')),
        temp %>% filter(group=='大豆/坚果 (g/d)'|group=='奶类 (g/d)') %>% mutate(min=c(25,300), 说明=case_when(摄入量<min~'摄入偏低，推荐适当增加摄入', TRUE~'摄入合理')),
          temp %>% filter(!(group=='油 (g/d)'|group=='大豆/坚果 (g/d)'|group=='蛋类 (g/d)'|group=='奶类 (g/d)')) %>% mutate(min=c(130,40,40,300,200,200,1500),max=c(180,65,65,500,300,250,1700),
                                                                                           说明=case_when(摄入量>max~'摄入偏高，推荐适当降低摄入', 摄入量<min~'摄入偏低，推荐适当增加摄入', TRUE~'摄入合理'))
          ) %>% transform(摄入量=round(摄入量,2)) %>%
         select(-min,-max) %>% relocate(备孕及孕早期孕妇推荐摄入量,.after = 说明) %>%
      mutate(group=factor(group, levels = rev(group))) %>% arrange(group) %>% rename('食物类别'='group')

    row.names(food_table) <- NULL

    food_format <- formattable(food_table, list(
      area(col = 摄入量) ~ normalize_bar("pink", 0.2),
      说明 = formatter("span",
                             style = x ~ ifelse(x == "摄入合理",style(color = "green"),
                                                            style(color = "red", font.weight = "bold"))))
    )

    nutrients <- rep(daily_quantify[-78,]%>% select(data[i,]$id),15)*component[,c(4,6:17,19,20)] %>% mutate(脂肪=(脂肪/能量)*100)

    nutrients <- nutrients %>% apply(2,function(x) sum(x,na.rm = TRUE)) %>% sapply(function(x) round(x,2)) %>% as.data.frame()

    temp1 <- nutrients %>% mutate(营养素=c('能量 (kcal/d)','蛋白质 (g/d)','脂肪 (g/d)','膳食纤维 (g/d)', '碳水化合物 (g/d)', '维生素A (μg/d)', '维生素B1 (mg/d)', '维生素B2 (mg/d)', '烟酸 (mg/d)', '维生素E (μg/d)', '钠 (mg/d)',
                                                  '钙 (mg/d)', '铁 (mg/d)', '维生素C (mg/d)', '胆固醇 (mg/d)'),
                                            min= c(1800, 50, 20, 25, 50, 700, 1.2,1.2, 10, 14, 2000, 800, 20, 100, 300),
                                            max= c(2400, 55, 30, 30, 65, 3000, 1.2,1.2, 12, 700, 2000, 2000, 42, 2000, 300),
                                            孕早期孕妇推荐摄入量= c('1800-2400kcal/d', '50-55 g/d', '20-30%占能量的百分比', '20-30g/d','50-65g/d', '700-3000μg/d', '>1.2mg/d','>1.2mg/d', '10-12mg/d', '12-15mg/d', '<2000mg/d', '800-2000mg/d','20-42mg/d','100-2000mg/d','<300mg/d' ))

    names(temp1)[1] <- '摄入量'

    nutrient_table <- temp1 %>% filter(营养素=='维生素B1 (mg/d)'|营养素=='维生素B2 (mg/d)') %>% mutate(说明=case_when(摄入量<min~'摄入偏低，推荐适当增加摄入', TRUE~'摄入合理')) %>% bind_rows(
      temp1 %>% filter(营养素=='钠 (mg/d)'|营养素=='胆固醇 (mg/d)') %>% mutate(说明=case_when(摄入量>min~'摄入偏高，推荐适当减少摄入', TRUE~'摄入合理')),
      temp1 %>% filter(!(营养素=='钠 (mg/d)'|营养素=='胆固醇 (mg/d)'|营养素=='维生素B1 (mg/d)'|营养素=='维生素B2 (mg/d)')) %>% mutate(说明=case_when(摄入量>max~'摄入偏高，推荐适当降低摄入', 摄入量<min~'摄入偏低，推荐适当增加摄入', TRUE~'摄入合理'))) %>%
        select(-min,-max) %>% relocate(孕早期孕妇推荐摄入量,.after = 说明) %>%
        relocate(营养素, .before = 摄入量) %>%
        mutate(营养素=factor(营养素, levels = c('能量 (kcal/d)','蛋白质 (g/d)','脂肪 (g/d)','膳食纤维 (g/d)', '碳水化合物 (g/d)', '维生素A (μg/d)', '维生素B1 (mg/d)', '维生素B2 (mg/d)', '烟酸 (mg/d)', '维生素E (mg/d)', '钠 (mg/d)',
                                          '钙 (mg/d)', '铁 (mg/d)', '维生素C (mg/d)', '胆固醇 (mg/d)'))) %>% arrange(营养素)

    row.names(nutrient_table) <- NULL

    nutrients_format <- formattable(nutrient_table, list(
      area(col = 摄入量) ~ normalize_bar("pink", 0.2),
      说明 = formatter("span",
                     style = x ~ ifelse(x == "摄入合理",style(color = "green"),
                                        style(color = "red", font.weight = "bold"))))
    )

    export_formattable <- function(f, file, width = "100%", height = NULL,
             background = "white", delay = 0.2)

    {
      w <- as.htmlwidget(f, width = width, height = height)
      path <- html_print(w, background = background, viewer = NULL)
      url <- paste0("file:///", gsub("\\\\", "/", normalizePath(path)))
      webshot(url,
              file = file,
              selector = ".formattable_widget",
              delay = delay)

    }

    export_formattable(food_format, file.path(tempdir(), 'food.png'))
    export_formattable(nutrients_format, file.path(tempdir(), 'nutrients.png'))

    img.file1 <- file.path(tempdir(), 'food.png')
    img.file2 <- file.path(tempdir(), 'nutrients.png')

    new_doc <- read_docx('Doc2.docx') %>%
      body_add_par(value = paste0('姓名: ',data[i,]$name,'                   ', '编号:', data[i,]$id), style = 'Style2') %>%
      body_add_par(value= "膳食结构分析与评价表", style = "Style1") %>%
      body_add_img(src = img.file1, height = 3.06, width = 7.39, style = 'Normal') %>%
      slip_in_footnote(blocks = block_list(
        fpar(ftext("本报告数据仅来源于膳食频率问卷，无临床指导意义",fp_text(font.size=8)))), pos = "before") %>%
      body_add_par(value= "能量和营养素摄入分析与评价表", style = "Style1") %>%
      body_add_img(src = img.file2, height = 4.06, width = 7.39, style = 'Normal')

    print(new_doc, target = file.path(tempdir(), "example_table.docx"))

    docx2pdf(input = file.path(tempdir(), "example_table.docx"), output = paste0(dir,data[i,]$id,'.pdf'))

  }



  }


ffq_report(data = '~/Desktop/Copy of 珠海队列.xlsx', dir = '~/Desktop/guzcs/')
