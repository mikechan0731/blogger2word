# -*-coding: utf-8-*-
# editor: MikeChan
# email: m7807031@gmail.com
# license: BSD-3

# OS import
import os
import random
import sqlite3
from datetime import date

# Parser import
import requests
from urllib import request
from bs4 import BeautifulSoup

# Office-docx lib import
import docx
from docx.shared import Cm, Pt  #加入可調整的 word 單位
from docx.enum.text import WD_ALIGN_PARAGRAPH #處理字串的置中

# ===== TODO ===== #
# 1.
# 2.
# 3.


# ===== Page Structure ===== #
# Blogger -> Year -> Month -> Page
# Page -> [Title, Date, Content]
# Content -> 


# ===== Global Variables ===== #
test_url_01 = "http://lulwechange.blogspot.com/2018/12/1.html" # with pic
test_url_02 = "http://lulwechange.blogspot.com/2018/10/blog-post_21.html" # with pics & font layout


# ====== functions ===== #

def blogger_to_year_list(blogger_url):
    pass


def year_list_to_page_list(year_url):
    pass


def single_page_to_content(url):
    target = url
    res = requests.get(target)
    if res.status_code == requests.codes.ok:
        soup = BeautifulSoup(res.text, "html.parser")

        # page basic info
        # title
        blogger_title = soup.title.text.split(":")[0].strip()
        page_title = soup.title.text.split(":")[1].strip()
        print(page_title)

        # date
        date_tag = soup.findAll("h2", class_ = "date-header")
        date_str = date_tag[0].text
        date ="".join([s for s in date_tag[0].text.split(" ")[0] if s.isdigit()])
        print(date)
        
        # tag
        tags_tag = soup.findAll(rel="tag")
        tags_list = []
        for i in range(len(tags_tag)):
            tags_list.append(tags_tag[i].text)
        print(tags_list)

        print("= = = = = = = = = =")

        # main content in page
        main_tag = soup.findAll(id="main")
        #main_text = main_tag[0].text


        # save to txt
        with open("tmp/tmp.txt", "w") as f:
            f.write(str(main_tag[0]))

        # read txt, tranfer, and save to docx
        docx_name = date + page_title + ".docx"
        txt_line_to_docx("tmp/tmp.txt", docx_name, page_title, date_str, tags_list)

        # Done
        print("Page: " + date + "_" + page_title + " -----Saved.")

    else:
        print("Get " + url + " Failed!!")


# read txt, tranfer, and save to docx
def txt_line_to_docx(txt_name, docx_name, title, date, tags_list):

    with open(txt_name, "r") as r:
    
        # check line style and write to docx
        doc = docx.Document()
        doc.add_heading(title, level=1)
        doc.add_heading(date, level=3)

        tags_str = ""
        for tag in tags_list:
            tags_str += "#" + tag + "; "
        doc.add_heading(tags_str, level=3)

        for line in r:

            # end of main content
            if u"張貼者" in line:
                break


            soup_line = BeautifulSoup(line, "lxml")
            

            # check if href in line
            if soup_line.findAll('a', href=True):
                text = soup_line.text.strip()
                url = soup_line.findAll('a', href=True)[0]['href']

                # img    
                if url.split(".")[-1] == "jpg" or url.split(".")[-1] == "png" or url.split(".")[-1] == "bmp":
                    request.urlretrieve(url,"tmp/tmp.jpg")
                    doc.add_picture("tmp/tmp.jpg", width=Cm(11))
                    last_paragraph = doc.paragraphs[-1] 
                    last_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER        

                # hyper-link
                else:
                    ok_line = text + "[" + url + "]"
                    doc.add_paragraph(ok_line)
            
            # normal text
            else:
                if soup_line.text.strip() == "":
                    continue 
                else:
                    if "<tr>" in line:
                        paragraph = doc.add_paragraph(soup_line.text)
                        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    else:
                         paragraph = doc.add_paragraph(soup_line.text)

        
        # save docx
        doc.save(docx_name)

# ====== main() ===== #
def main():
    if not os.path.exists('tmp'):
        os.makedirs('tmp')

    single_page_to_content(test_url_02)

    print("** Done! **")


if __name__=="__main__":
    main()