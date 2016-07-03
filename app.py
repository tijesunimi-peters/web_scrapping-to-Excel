from openpyxl import *
from bs4 import *
from urllib2 import *
from urllib import *
from re import *
from requests import *

wb = load_workbook(filename = "Mass Job Upload.xlsx")
active = wb["Fields"]
def state():
  state = raw_input("State: ")
  return state.lower().replace(" ","-")

def page():
  start = raw_input("Starting Page: ")
  limit = raw_input("Page Limit: ")
  return {
    "start": int(start), "limit": int(limit)
  }

def per_page():
  perpage = raw_input("Jobs Per Page: ")
  return int(perpage)

def v():
  return {
    "state": state(),
    "page": page(),
    "per-page": per_page()
  }

var = v()

start_row = 2
col = 2

def getSite(page):
  url = "http://www.careers24.com.ng/jobs/lc-%s/m-true/?sort=dateposted&pagesize=%d&page=%d" % (var["state"], var["per-page"], page)
  return get(url)

def getLinks(site):
  search_links = BeautifulSoup(site.text, "html.parser")
  return search_links.select("h4 > a")


for page in range(var["page"]["start"],var["page"]["limit"]):

  print "\nPage: ",page
  site = getSite(page)
  if site.status_code == codes.ok:
    link_count = 0
    for link in getLinks(site):
      link_count += 1
      print "Link: ", link_count,
      url = 'http://www.careers24.com.ng'+link.attrs['href']
      page_url = get(url)
      page_content = BeautifulSoup(page_url.text, "html.parser")
      active.cell(column=col, row = 1, value=var["state"])
      active.cell(column=col, row = 20, value="%s" % page_content.select('[itemprop="title"]')[0].getText())
      print page_content.select('[itemprop="title"]')[0].getText()
      active.cell(column=col, row = 7, value="%s" % page_content.select('[itemprop="baseSalary"]')[0].getText())
      
      description = page_content.select('[itemprop="hiringOrganization"] > ul')
      if len(description) > 0:
        active.cell(column=col, row = 21).value = ""
        for item in description:
          try:
            active.cell(column=col, row = 21).value += escape(item.getText())
          except Exception, e:
            pass
          

      requirements = page_content.select('[itemprop="responsibilities"] > ul')
      if len(requirements) > 0:
        active.cell(column=col, row = 23).value = ""
        for item in requirements:
          active.cell(column=col, row = 23).value += item.getText()
      active.cell(column=col, row = 24, value=url)
      col += 1
  else:
    print "Page Not Found"

  output_name = "%s-%d.xlsx" % (var["state"].capitalize(), var["page"]["start"])
  wb.save(output_name)

