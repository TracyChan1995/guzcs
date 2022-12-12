pa_report <- function(data = data, dir= dir) {

##体力活动及睡眠报告


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

if (suppressPackageStartupMessages(!require(readxl)))  {
  install.packages('readxl')
  suppressPackageStartupMessages(library(readxl))
}

data <- read_dta(file = data)

pa <- data %>% select(starts_with('p'))

sleep <- data %>% select(id,name,starts_with('s'))

主观睡眠质量 <- sleep %>% select(id,name,s6s1, s6s2) %>% mutate(孕前一年得分=case_when(s6s1==1~'0', s6s1==2~'1', s6s1==3~'2', TRUE~'3'),
                                                     最近一月得分=case_when(s6s2==1~'0', s6s2==2~'1', s6s2==3~'2', TRUE~'3'),
                                                     成分='主观睡眠质量',孕前一年得分=as.numeric(孕前一年得分), 最近一月得分=as.numeric(最近一月得分)) %>% select(-s6s1, -s6s2)

入睡时间 <- sleep %>% select(id,name,s2s1,s2s2, s5as1, s5as2) %>% mutate(score1=case_when(s2s1<=15~'0', s2s1>=16 & s2s1<=30~'1', s2s1>=31 & s2s1<=60~'2', TRUE~'3'),
                                                           score2=case_when(s2s2<=15~'0', s2s2>=16 & s2s2<=30~'1', s2s2>=31 & s2s2<=60~'2', TRUE~'3'),
                                                           score3=case_when(s5as1==1~'0', s5as1==2~'1', s5as1==3~'2', TRUE~'3'),
                                                           score4=case_when(s5as2==1~'0', s5as2==2~'1', s5as2==3~'2', TRUE~'3'),
                                                           score1= as.numeric(score1), score2 = as.numeric(score2), score3 = as.numeric(score3), score4 = as.numeric(score4),
                                                           score5=score1+score3, score6= score2+ score4) %>%
                                                    mutate(孕前一年得分=case_when(score5==0~'0', score5>=1 & score5<=2~'1', score5>=3 & score5<=4~'2', TRUE~'3'),
                                                           最近一月得分=case_when(score6==0~'0', score6>=1 & score6<=2~'1', score6>=3 & score6<=4~'2', TRUE~'3'),
                                                           成分='睡眠时间',孕前一年得分=as.numeric(孕前一年得分), 最近一月得分=as.numeric(最近一月得分)) %>% select(-s2s1,-s2s2, -s5as1, -s5as2, -score1, -score2, -score3, -score4, -score5, -score6)

睡眠时间 <- sleep %>% select(id,name,s4s1, s4s2) %>% mutate(孕前一年得分=case_when(s4s1>7~'0',s4s1>=6 & s4s1<=7~'1', s4s1>=5 & s4s1<=6~'2', TRUE~'3'),
                                                  最近一月得分=case_when(s4s2>7~'0',s4s2>=6 & s4s2<=7~'1', s4s2>=5 & s4s2<=6~'2', TRUE~'3'),
                                                  孕前一年得分=as.numeric(孕前一年得分), 最近一月得分=as.numeric(最近一月得分),
                                                  成分='睡眠时间') %>%
                                           select(-s4s1, -s4s2)

睡眠效率 <- sleep %>% select(id,name,s4s1, s4s2, s3s1, s3s2, s3s3,s3s4, s1s1, s1s2, s1s3, s1s4) %>% mutate(score1= ifelse(s3s1>s1s1, s3s1+(s3s2/60)-s1s1-(s1s2/60), s3s1+(s3s2/60)+(24-(s1s1+(s1s2/60)))),
                                                                                                   score2= ifelse(s3s3>s1s3, s3s3+(s3s4/60)-s1s3-(s1s4/60), s3s3+(s3s4/60)+(24-(s1s3+(s1s4/60)))),
                                                                                                   score3= s4s1/score1, score4= s4s2/score2,
                                                                                               孕前一年得分=case_when(score3>0.85~'0', score3<=0.85 & score3>0.75~'1', score3<=0.75 & score3>0.65~'2', TRUE~'3'),
                                                                                               最近一月得分=case_when(score4>0.85~'0', score4<=0.85 & score4>0.75~'1', score4<=0.75 & score4>0.65~'2', TRUE~'3')) %>%
                                                    mutate(孕前一年得分=as.numeric(孕前一年得分), 最近一月得分=as.numeric(最近一月得分), 成分='睡眠效率') %>% select(-score1, -score2, -score3, -score4, -s4s1, -s4s2, -s3s1, -s3s2, -s3s3, -s3s4, -s1s1, -s1s2, -s1s3, -s1s4)


睡眠障碍 <- sleep %>% select(id,name,s5bs1, s5bs2, s5cs1, s5cs2, s5ds1, s5ds2, s5es1, s5es2,s5fs1, s5fs2,s5gs1, s5gs2,s5hs1, s5hs2,s5is1, s5is2,s5ks1, s5ks2) %>%
                      select(ends_with('s1')) %>% apply(1, sum) %>% as_tibble() %>%
                      mutate(孕前一年得分=case_when(value==0~'0', value>=1 & value<=9 ~'1', value>=10 & value<= 18~'2', TRUE~'3')) %>% select(-value) %>% bind_cols(
                        sleep %>% select(s5bs1, s5bs2, s5cs1, s5cs2, s5ds1, s5ds2, s5es1, s5es2,s5fs1, s5fs2,s5gs1, s5gs2,s5hs1, s5hs2,s5is1, s5is2,s5ks1, s5ks2) %>%
                          select(ends_with('s2')) %>% apply(1,  sum) %>% as_tibble() %>%
                          mutate(最近一月得分=case_when(value==0~'0', value>=1 & value<=9 ~'1', value>=10 & value<= 18~'2', TRUE~'3')) %>% select(-value)
                      ) %>% mutate(成分='睡眠障碍', 孕前一年得分=as.numeric(孕前一年得分), 最近一月得分=as.numeric(最近一月得分)) %>% bind_cols(sleep %>% select(id,name))

催眠药物 <- sleep %>% select(id,name,s7s1, s7s2) %>% mutate(孕前一年得分=case_when(s7s1==1~'0', s7s1==2~'1', s7s1==3~'2', TRUE~'3'),
                                                    最近一月得分=case_when(s7s2==1~'0', s7s2==2~'1', s7s2==3~'2', TRUE~'3')) %>%
                                             mutate(孕前一年得分=as.numeric(孕前一年得分), 最近一月得分=as.numeric(最近一月得分), 成分='催眠药物') %>%
                                             select(-s7s1, -s7s2)

日间功能障碍 <- sleep %>% select(id,name,s8s1, s8s2, s9s1, s9s2) %>% mutate(score1=s8s1+s9s1, score2= s8s2+s9s2,
                                                                    孕前一年得分= case_when(score1 ==0~'0', score1>=1 & score1 <=2~'1', score1>=3 & score1<=4~'2', TRUE~'3'),
                                                                    最近一月得分= case_when(score2 ==0~'0', score2>=1 & score2 <=2~'1', score2>=3 & score2<=4~'2', TRUE~'3')) %>%
                                                             mutate(孕前一年得分=as.numeric(孕前一年得分), 最近一月得分=as.numeric(最近一月得分), 成分='日间功能障碍') %>%
                                                             select(-s8s1, -s8s2, -s9s1, -s9s2, -score1, -score2)

sleep_data <- 主观睡眠质量 %>% bind_rows(入睡时间, 睡眠时间, 睡眠效率, 睡眠障碍, 催眠药物, 日间功能障碍)

total_sleep <- sleep_data %>% group_by(id,name) %>% summarise(孕前一年睡眠质量总得分=sum(孕前一年得分),
                                                最近一月睡眠质量总得分=sum(最近一月得分)) %>% mutate(孕前一年得分=ifelse(孕前一年睡眠质量总得分>5, '睡眠质量偏低', '睡眠质量良好'), 最近一月得分=ifelse(最近一月睡眠质量总得分>5, '睡眠质量偏低', '睡眠质量良好'))

##pa

pa <- data %>% select(starts_with('p')) %>% select(p1p1:p8p2,p11p1:p28p2p2) %>% mutate_at(vars(starts_with('p')),
                                                                                    list(~case_when(
                                                                                      . == "1"  ~ 0,
                                                                                      . == '2' ~ 0.25,
                                                                                      . == "3" ~ 0.75,
                                                                                      . == "4" ~ 1.5,
                                                                                      . == "5" ~ 2.5,
                                                                                      . == "6" ~ 3))) %>% bind_cols(
data %>% select(starts_with('p')) %>% select(p9p1, p9p2, p10p1, p10p2, p29p1:p33p2) %>%
         mutate_at(vars(starts_with('p')),
                       list(~case_when(
                         . == "1"  ~ 0,
                         . == '2' ~ 0.25,
                         . == "3" ~ 1.25,
                         . == "4" ~ 3,
                         . == "5" ~ 5,
                         . == "6" ~ 6)))) %>% relocate(p9p1:p10p2, .after = p8p2)

pre_data <- pa %>% select(ends_with('p1')) %>% select(-(starts_with('p28')))

early_data <- pa %>% select(ends_with('p2')) %>% select(-(starts_with('p28')))

pre_data <- pre_data*7*met$MET

early_data <- early_data*7*met$MET

pre_total <- pre_data %>% apply(1, sum) %>% as.tibble() %>%  mutate(体力类型='总体力活动 (MET-h/week)') %>%
                      bind_rows(
                      pre_data %>% select(starts_with('p8')|starts_with('p9')|starts_with('p10')|starts_with('p29')) %>% apply(1, sum) %>% as.tibble() %>% mutate(体力类型='久坐活动 (MET-h/week)'),
                      pre_data %>% select(starts_with('p1')|starts_with('p2')|starts_with('p4')|starts_with('p12')|starts_with('p13')|starts_with('p14')|starts_with('p15')|starts_with('p17')|starts_with('p31')) %>% apply(1, sum) %>% as.tibble() %>% mutate(体力类型='轻强度体力活动 (MET-h/week)'),
                      pre_data %>% select(starts_with('p3')|starts_with('p5')|starts_with('p6')|starts_with('p7')|starts_with('p11')|starts_with('p16')|starts_with('p18')|starts_with('p20')|starts_with('p21')|starts_with('p24')|starts_with('p25')|starts_with('p26')|starts_with('p30')|starts_with('p32')|starts_with('p33')) %>% apply(1, sum) %>% as.tibble() %>% mutate(体力类型='中强度体力活动 (MET-h/week)'),
                      pre_data %>% select(starts_with('p22')|starts_with('p23')) %>% apply(1, sum) %>% as.tibble() %>% mutate(体力类型='高强度体力活动 (MET-h/week)'),
                      pre_data %>% apply(1, sum) %>% as.tibble() %>%  mutate(体力类型='说明', value= ifelse(value<7.5, '体力活动不足，建议适当增加中等强度的体力活动', '体力活动充足'))
                      ) %>% rename(孕前一年='value')

early_total <- early_data %>% apply(1, sum) %>% as.tibble() %>%  mutate(体力类型='总体力活动 (MET-h/week)') %>%
  bind_rows(
    early_data %>% select(starts_with('p8')|starts_with('p9')|starts_with('p10')|starts_with('p29')) %>% apply(1, sum) %>% as.tibble() %>% mutate(体力类型='久坐活动 (MET-h/week)'),
    early_data %>% select(starts_with('p1')|starts_with('p2')|starts_with('p4')|starts_with('p12')|starts_with('p13')|starts_with('p14')|starts_with('p15')|starts_with('p17')|starts_with('p31')) %>% apply(1, sum) %>% as.tibble() %>% mutate(体力类型='轻强度体力活动 (MET-h/week)'),
    early_data %>% select(starts_with('p3')|starts_with('p5')|starts_with('p6')|starts_with('p7')|starts_with('p11')|starts_with('p16')|starts_with('p18')|starts_with('p20')|starts_with('p21')|starts_with('p24')|starts_with('p25')|starts_with('p26')|starts_with('p30')|starts_with('p32')|starts_with('p33')) %>% apply(1, sum) %>% as.tibble() %>% mutate(体力类型='中强度体力活动 (MET-h/week)'),
    early_data %>% select(starts_with('p22')|starts_with('p23')) %>% apply(1, sum) %>% as.tibble() %>% mutate(体力类型='高强度体力活动 (MET-h/week)'),
    early_data %>% apply(1, sum) %>% as.tibble() %>%  mutate(体力类型='说明', value= ifelse(value<7.5, '体力活动不足，建议适当增加中等强度的体力活动', '体力活动合理'))
  ) %>% rename(最近一月='value')

data_report <- data.frame(id=rep(data$id,6)) %>% bind_cols(pre_total %>% select(-体力类型),early_total) %>% relocate(体力类型, .after = id)

for (i in 1:length(data$id)) {
  temp1 <- data_report %>% filter(id==data[i,]$id) %>% select(-id)
  temp2 <- sleep_data %>%  filter(id==data[i,]$id) %>% bind_rows(total_sleep %>% filter(id==data[i,]$id) %>% select(孕前一年睡眠质量总得分, 最近一月睡眠质量总得分) %>% rename(孕前一年得分='孕前一年睡眠质量总得分', 最近一月得分='最近一月睡眠质量总得分') %>% mutate(成分='总得分')) %>% select(-id, -name) %>% relocate(成分,.before = 孕前一年得分)

  pa_format <- formattable(temp1, list(
    area(col = c(孕前一年, 最近一月)) ~ normalize_bar("pink", 0.6),
    体力活动 = formatter("span",
                   style = x ~ ifelse(x == "体力活动合理",style(color = "green"),
                                      style(color = "red", font.weight = "bold"))))
  )

  sleep_format <- formattable(temp2, list(
    area(col = c(孕前一年得分, 最近一月得分)) ~ normalize_bar("pink", 0.6),
    成分 = formatter("span",
                     style = x ~ ifelse(x != "总得分",style(color = "green"),
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

  export_formattable(pa_format, file.path(tempdir(), 'pa.png'))
  export_formattable(sleep_format, file.path(tempdir(), 'sleep.png'))

  pa_img <- file.path(tempdir(), 'pa.png')
  sleep_img <- file.path(tempdir(), 'sleep.png')

  new_doc <- docx %>%
    body_add_par(value = paste0('姓名: ',data[i,]$name,'                   ', '编号:', data[i,]$id), style = 'Style2') %>%
    body_add_par(value= "体力活动分析与评价表", style = "Style1") %>%
    body_add_img(src = pa_img, height = 3.06, width = 7.39, style = 'Normal') %>%
    slip_in_footnote(blocks = block_list(
      fpar(ftext("本报告数据仅来源于孕期体力活动量表，无临床指导意义",fp_text(font.size=8)))), pos = "before") %>%
    body_add(fpar(run_linebreak())) %>%
    body_add_img(src = pa_tip, height = 3.06, width = 5.39, style = 'Style1') %>%
    body_add(fpar(run_linebreak())) %>%
    body_add(fpar(run_linebreak())) %>%
    body_add(fpar(run_linebreak())) %>%
    body_add(fpar(run_linebreak())) %>%
    body_add_par(value= "睡眠质量分析与评价表", style = "Style1") %>%
    body_add_img(src = sleep_img, height = 3.06, width = 7.39, style = 'Normal') %>%
    slip_in_footnote(blocks = block_list(
      fpar(ftext("本报告数据仅来源于匹兹堡睡眠质量量表，无临床指导意义",fp_text(font.size=8)))), pos = "before") %>%
    body_add(fpar(run_linebreak())) %>%
    body_add_img(src = sleep_tip, height = 3.06, width = 5.39, style = 'Style1')


  print(new_doc, target = "example_table.docx")

  docx2pdf(input = file.path(tempdir(), "example_table.docx"), output = paste0(dir,data[i,]$id,'.pdf'))

}
}
