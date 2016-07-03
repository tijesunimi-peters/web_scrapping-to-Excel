from openpyxl import load_workbook
from bs4 import BeautifulSoup
from re import escape
from requests import get, codes, adapters
from time import sleep

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

def getSite(page):
  url = "http://www.careers24.com.ng/jobs/lc-%s/m-true/?sort=dateposted&pagesize=%d&page=%d" % (var["state"], var["per-page"], page)
  return get(url)

def getLinks(site):
  search_links = BeautifulSoup(site.text, "html.parser")
  return search_links.select("h4 > a")


for page in range(var["page"]["start"],var["page"]["limit"]):
  col = 2
  print "\nPage: ",page
  output_name = "%s-%d.xlsx" % (var["state"].capitalize(), page)
  try:
    site = getSite(page)
    if site.status_code == codes.ok:
      link_count = 0
      for link in getLinks(site):
        link_count += 1
        print "Link: ", link_count,
        url = 'http://www.careers24.com.ng'+link.attrs['href']

        try:
          page_url = get(url)
        except Exception, e:
          print("Connection Failure\nSaving File and Aborting")
          wb.save(output_name)

        
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
        print("sleeping for 5 secs")
        sleep(5)
    else:
      print "Page Not Found"

    wb.save(output_name)

  except KeyboardInterrupt:
    print "Exiting......\n Previous work saved"
    wb.save(output_name)
  

  
  

