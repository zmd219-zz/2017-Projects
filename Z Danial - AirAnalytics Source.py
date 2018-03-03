"""
Zach Danial

This documennt contains the source code for an algorithm I created in the summer of 2017 to webscrape Airbnb
listings and instntly produce insights into the short term rental market.  Due to updates in the
html structure of the Airbnb website the scraping algorithm no longer functions.  The code
generates a chart that details the distribution of prices by bed-count for a given neighborhood or
city.  It additionally generates a 30-day vacancy report that allows hosts to identify high demand 
days.  Finally, it charts the distribution of prices by bed-count for each particular high traffic 
day, allowing hosts to understand how the market is pricing these specific days.

To see an exampled of the graphs outputted by this project please click the link in the file titled
Air Analytics Dashboard Example in my portfolio.

"""


#import requirements
from lxml import html
import datetime
import requests
import plotly.tools
import plotly.plotly as py
import plotly.graph_objs as go
import plotly.figure_factory as ff
from scipy.signal import argrelextrema
import numpy as np
import pandas as pd
import colorlover as cl
from openpyxl import load_workbook



start_time = datetime.datetime.now()

#configuring plot package
plotly.tools.set_config_file(world_readable=False,
                             sharing='private')
scheme = 5*cl.scales['9']['qual']['Pastel1']

#counters to track numbers of listings grabbed
total_accuracy = 0
accuracy_count = 0


#This class allows chopping a coordinate grid into smaller peices 
#to ensure all listings in the given area are retreived and allows 
#generation of web adresses for iteration over chopped grid
class MapSearch:
    def __init__(self, base, sw_lat, sw_lng, ne_lat, ne_lng):
        self.base = base
        self.ne_lat = ne_lat
        self.ne_lng = ne_lng
        self.sw_lat = sw_lat
        self.sw_lng = sw_lng
        self.area_total_price = 0
        self.area_total_units = 0

    def add_units(self, increment):
        self.area_total_units += increment

    def add_price(self, price):
        self.area_total_price += price

    def area_total_average(self):
        if self.area_total_units > 0:
            out = self.area_total_price/self.area_total_units
        else:
            out = None
        return out

    def query(self):
        return str(self.base + "&ne_lat=" + str(self.ne_lat) + "&ne_lng=" + str(self.ne_lng) + "&sw_lat=" +
                   str(self.sw_lat) + "&sw_lng=" + str(self.sw_lng) + "&search_by_map=true")

    def chop(self, pieces):
        lng_start = self.sw_lng
        lng_step = abs((self.ne_lng - self.sw_lng) / pieces)
        lng_end = self.ne_lng

        lat_start = self.sw_lat
        lat_step = abs((self.sw_lat - self.ne_lat) / pieces)
        lat_end = self.ne_lat
        print(lat_start, lat_end, lat_step, (lat_end-lat_start)/lat_step)
        print(lng_start, lng_end, lng_step, (lng_end-lng_start)/lng_step)

        start = np.asarray([lat_start, lng_start, lat_start + lat_step, lng_start + lng_step])
        box_on = start
        out = []
        for i in range(pieces):
            for x in range(pieces):
                out.append(MapSearch(self.base, box_on[0]+x*lat_step, box_on[1], box_on[2]+x*lat_step, box_on[3]))
                print(box_on[0]+x*lat_step, box_on[1], box_on[2]+x*lat_step, box_on[3])
            box_on = start
            box_on[1] += lng_step
            box_on[3] += lng_step
        print()
        return out

    @staticmethod
    def clean(base, string):
        first = string.split("&")
        for item in first:
            if "ne_lat=" in item:
                start = item.find("ne_lat=")
                ne_lat = float(item[start+7:])
            elif "ne_lng=" in item:
                start = item.find("ne_lng=")
                ne_lng = float(item[start+7:])
            elif "sw_lng=" in item:
                start = item.find("sw_lng=")
                sw_lng = float(item[start+7:])
            elif "sw_lat=" in item:
                start = item.find("sw_lat=")
                sw_lat = float(item[start+7:])
        return MapSearch(base, sw_lat, sw_lng, ne_lat, ne_lng)



#Date class is an object for storage of data from scraping specific days 
class Date:
    calendar = []
    scan = []

    def __init__(self, text):
        if "/" in text:
            self.month = text[:2]
            self.string_day = text[3:5]
            self.day = int(text[3:5])
            self.year = text[6:]
        else:
            self.month = text[5:7]
            self.string_day = text[-2:]
            self.day = int(self.string_day)
            self.year = text[:4]

    def out(self):
        return self.year + "-" + self.month + "-" + add_zero(self.day)

    def add_days(self, amount):
        length = month_length(self)
        month = self.month
        day = self.day + amount
        year = self.year
        if day > length:
            day -= length
            month = add_zero(int(month)+1)
        if int(month) > 12:
            year = add_zero(int(year) + 1)
            month = add_zero(int(month)-12)
        return Date(str(year)+"-"+str(month)+"-"+str(add_zero(day)))

    def path(self):
        start = self
        end = Date.add_days(self, 1)
        return '&checkin=' + start.out() + '&checkout=' + end.out()

    @staticmethod
    def _info(term):
        for item in Date.calendar:
            if item.day == term:
                return item.out()

    @staticmethod
    def list_out(group):
        out = []
        for item in group:
            out.append(item.out())
        return out

    @staticmethod
    def to_form(date):
        month = add_zero(date.month)
        day = add_zero(date.day)
        year = str(date.year)
        text = year + "-" + month + "-" + day
        return Date(text)


#Listing object store information of a particular listings from a non-specific date scrape
#Represents minimum price for a listing (baseline price)
class Listing:
    _baseline = []
    all_ids = []
    highest_bed_number = 0
    highest_price = 0

    def __init__(self, listing_id, title, typo, city, price, beds, rating, review_count):
        self.id = listing_id
        Listing.all_ids.append(listing_id)
        self.typo = typo #type
        self.city = city
        self.price = price
        self.beds = beds
        self.title = title
        self.rating = rating
        self.review_count = review_count
        Listing._baseline.append(self)
        if Listing.highest_bed_number < beds:
            Listing.highest_bed_number = beds
        if Listing.highest_price < price:
            Listing.highest_price = price

    def out(self):
        return self.id, self.title, self.typo, self.city, self.price, self.beds, self.rating, self.review_count

    @classmethod
    def list_by_bed(cls, bed):
        temp_list = []
        scan = cls._baseline
        for listing in scan:
            if listing.beds == bed:
                temp_list.append(listing.price)
        return temp_list

    @classmethod
    def count_by_bed(cls, bed):
        count = 0
        scan = cls._baseline
        for listing in scan:
            if listing.beds == bed:
                count += 1
        return count

    @classmethod
    def search_by_id(cls, ident):
        for item in Listing._baseline:
            if item.id == ident:
                return item

    @classmethod
    def search_by_title(cls, title):
        for item in Listing._baseline:
            if item.title == title:
                return item
    
    
    #This method sifts through the html code to gather the data for a specific listing
    #via the Xpath to the data values 
    #The reason the code no longer works is because as Airbnb updates their website, these
    #Xpaths change and have to be updated
    @staticmethod
    def listing_process(item):
        unit_info = {'id': '--', 'name': '--', 'typo': '--', 'city': '--', 'price': '--', 'beds': '--', 'rating': '--',
                     'review_count': '--', 'error': 0}
        try:
            name_type_city = item.xpath('meta[@itemprop="name"]/@content')[0].split(" - ")
        except IndexError:
            print('Listing Error - IndexError: 1')
            return [1, 0, 0, 0, 0, 0, 0, 0]
        else:
            if len(name_type_city) == 3:
                name = name_type_city[0]
            elif len(name_type_city) > 3:
                name = name_type_city[:-2]
                temp = ''
                i = 0
                for part in name:
                    temp += part
                    i += 1
                    if i < len(name):
                        temp += " - "
                name = temp
            else:
                print('Listing Error - NameError')
                unit_info['error'] = 1
                return [1, 0, 0, 0, 0, 0, 0, 0]
            typo = name_type_city[-2]
            city = name_type_city[-1]
            unit_info['name'] = name
            unit_info['typo'] = typo
            unit_info['city'] = city
        try:
            listing_card = item.xpath('div[@class="listingCardWrapper_9kg52c"]/div[@class="listingContainer_'
                                      'f21qs6"]')[0]
        except IndexError:
            print('Listing Error - IndexError: 2')
            print(item)
            unit_info['error'] = 1
            return [1, 0, 0, 0, 0, 0, 0, 0]
        except AttributeError:
            print('AttributeError - this was passed', item)
            return [1, 0, 0, 0, 0, 0, 0, 0]
        else:
            listing_id = listing_card.xpath('@id')[0][8:]
            unit_info['id'] = listing_id
            info_container = listing_card.xpath('div[@class="infoContainer_v72lrv"]')[0]
            raw = info_container.xpath('descendant::text()')

        try:
            rating_container = info_container.xpath('descendant::div[@class="ratingContainer_inline_36rlri"]')[0]
            rating_text = rating_container.xpath('descendant::span[@role="img"]/@aria-label')[0]
            rating_texts = rating_text.split(" ")
            rating = float(rating_texts[1])
            review_count = int(rating_container.xpath(
                'descendant::span[@class="text_5mbkop-o_O-size_micro_16wifzf-o_O-inline_g86r3e"]/text()')[0])
            unit_info['rating'] = rating
            unit_info['review_count'] = review_count
        except IndexError:
            unit_info['error'] = 1
            for bit in raw:
                if ' review' in bit:
                    stop = bit.find(' review')
                    if bit[:stop].isdigit():
                        num = int(bit[:stop])
                        unit_info['review_count'] = num
                        unit_info['error'] = 0
        for bit in raw:
            if bit[0] == '$' and (bit[1:].isdigit() or or_comma(bit[1:])):
                price = big_int(bit[1:])
                unit_info['price'] = price
            elif (bit[-4:] == ' bed' or bit[-5:] == ' beds') and (bit[:-4].isdigit() or bit[:-5].isdigit()):
                stop = bit.find(' bed')
                beds = big_int(bit[:stop])
                unit_info['beds'] = beds
        if unit_info['price'] == '--' or unit_info['beds'] == '--':
            print(raw)
        return unit_info
    
    
    #This method iterates over a page, and calls the listing_process method on
    #each listing on each page
    @staticmethod
    def baseline_processor(listings):
        page_total = len(listings)
        page_counter = 0
        total_prices = 0
        page_units = []
        errors = 0
        for item in range(0, len(listings)):
            if type(listings[item]) is str:
                print(listings)
            unit_info = Listing.listing_process(listings[item])
            if unit_info == [1, 0, 0, 0, 0, 0, 0, 0]:
                errors += 1
            if not unit_info == [1, 0, 0, 0, 0, 0, 0, 0]:
                listing_id = unit_info['id']
                name = unit_info['name']
                typo = unit_info['typo']
                city = unit_info['city']
                price = unit_info['price']
                beds = unit_info['beds']
                rating = unit_info['rating']
                review_count = unit_info['review_count']
                errors += unit_info['error']
                total_prices += price
                temp = Listing(listing_id, name, typo, city, price, beds, rating, review_count)
                page_units.append(temp)
                page_counter += 1
        if page_counter/page_total <= 0.8:
            for item in page_units:
                del item
            print('--retrying query--')
            return [0, 0]
        else:
            return [page_counter, total_prices]
        
        
    #This method iterates over the pages in a query, and calls the baseline_processor method
    #on each page
    @staticmethod
    def pages_iterate(pages, search, path):
        search_sample = 0
        search_total = 0
        search_on = search.query()
        for x in range(0, pages):
            print('Page: ' + str(x + 1), "/", str(pages))
            print(search.query())
            i = 0
            z = 0
            while True:
                if x == 0:
                    tree = grab(search_on)
                else:
                    tree = grab(search_on + path + str(x))
                listings = tree.xpath('//div[@itemprop="itemListElement"]')
                if len(listings) == 0:
                    i += 1
                    if i == 3:
                        print("N/A")
                        break
                    else:
                        continue
                else:
                    if type(listings) is list:
                        base_on = Listing.baseline_processor(listings)
                    else:
                        continue
                    if base_on == [0, 0]:
                        z += 1
                        if z == 2:
                            break
                        continue
                    else:
                        page_counter = base_on[0]
                        search_total += page_counter
                        search_sample += len(listings)
                        total_prices = base_on[1]
                        search.add_units(page_counter)
                        search.add_price(total_prices)
                        print(str(page_counter) + "/" + str(len(listings)) + "\n")
                        break
        return [search_total, search_sample]


    #This method combines the prior methods and creates a database for the baseline prices of
    #all listings in a given area
    #Output is accuracy metrics
    @staticmethod
    def baseline(path, searches):
        total_sample_total = 0
        total_total = 0
        search_num = 1
        for search_on in searches:
            print('Search: ' + str(search_num) + '/' + str(len(searches)))
            search_num += 1
            pages = get_pages(search_on.query())
            if pages >= 18:
                print('Baseline- ALERT: Listings missed - Search span included more than 18 pages')
            elif pages == 0:
                pages += 1
            totals = Listing.pages_iterate(pages, search_on, path)
            total = totals[0]
            sample_total = totals[1]
            if not sample_total == 0:
                total_total += total
                accuracy = 100 * total / sample_total
                global total_accuracy
                total_accuracy += accuracy
                global accuracy_count
                accuracy_count += 1
                total_sample_total += sample_total
                print('Sample of ' + str(total) + '/' + str(sample_total) + ' listings ('
                      + format(accuracy, '.2f') + '% accuracy)\n')
            else:
                print('Search provided no results.\n')
        if total_sample_total != 0:
            print('Sample of ' + str(total_total) + '/' + str(total_sample_total) + ' listings (' + format(
                100 * total_total / total_sample_total,
                '.2f') + '% accuracy)\n')
        return [total_total, total_sample_total]



#The ListingOnline class is almost exactly the same as the Listing class, however it represents
#a listing that is posted on a specific night
#'Online' terminology refers to being online on a specific night
class ListingOnline:
    in_the_month = []
    all_pulled = []

    def __init__(self, listing_id, date, price):
        self.id = listing_id
        ListingOnline.in_the_month.append(listing_id)
        self.vacant_dates = dict()
        self.vacant_dates[date.out()] = price
        self.info = Listing.search_by_id(listing_id)
        ListingOnline.all_pulled.append(self)

    def out(self):
        info = self.info.out()
        out = pd.Series(self.vacant_dates)
        return info, out

    @classmethod
    def search_by_id(cls, listing_id):
        for item in ListingOnline.all_pulled:
            if item.id == listing_id:
                return item

    @classmethod
    def list_by_bed_by_date(cls, bed, date):
        temp_list = []
        for listing in cls.all_pulled:
            if listing.info.beds == bed and date.out() in Date.list_out(listing.vacant_dates):
                temp_list.append(listing.vacant_dates[date.out()])
        return temp_list

    @classmethod
    def count_by_bed_by_date(cls, bed, date):
        count = 0
        for listing in cls.all_pulled:
            if listing.info.beds == bed and date.out() in Date.list_out(listing.vacant_dates):
                count += 1
        return count

    @staticmethod
    def vacant_date(listing_id, date, price):
        listing = ListingOnline.search_by_id(listing_id)
        listing.vacant_dates[date.out()] = price

    @staticmethod
    def date_price_process(item):
        unit_info = {'id': '--', 'price': '--'}
        try:
            listing_card = item.xpath('div[@class="listingCardWrapper_9kg52c"]/div[@class="listingContainer_'
                                      'f21qs6"]')[0]
        except IndexError:
            print('IndexError - Listing card')
            return [0, 0]
        else:
            listing_id = listing_card.xpath('@id')[0][8:]
            unit_info['id'] = listing_id
            info_container = listing_card.xpath('div[@class="infoContainer_v72lrv"]')[0]
            raw = info_container.xpath('descendant::text()')
            for bit in raw:
                if bit[0] == '$' and (bit[1:].isdigit() or or_comma(bit[1:])):
                    price = big_int(bit[1:])
                    unit_info['price'] = price
                    break
            if unit_info['price'] == '--':
                print(listing_id, raw)
                return [0, 0]
            else:
                if listing_id in Listing.all_ids:
                    return unit_info
                else:
                    unit_info = Listing.listing_process(item)
                    if not unit_info == [1, 0, 0, 0, 0, 0, 0, 0]:
                        listing_id = unit_info['id']
                        name = unit_info['name']
                        typo = unit_info['typo']
                        city = unit_info['city']
                        price = unit_info['price']
                        beds = unit_info['beds']
                        rating = unit_info['rating']
                        review_count = unit_info['review_count']
                        Listing(listing_id, name, typo, city, price, beds, rating, review_count)
                    else:
                        print('ListingOnline - Error')
                        return [0, 0]
                    return unit_info

    @staticmethod
    def page_processor(listings, date_on):
        page_counter = 0
        for item in listings:
            unit_info = ListingOnline.date_price_process(item)
            if not unit_info == [0, 0]:
                listing_id = unit_info['id']
                price = unit_info['price']
                if listing_id in ListingOnline.in_the_month:
                    ListingOnline.vacant_date(listing_id, date_on, price)
                else:
                    ListingOnline(listing_id, date_on, price)
                page_counter += 1
        if page_counter == 0 and len(listings) != 0:
            print('Page Error')
            return int(-5)
        return page_counter

    @staticmethod
    def online_pages_iterate(date_on, pages, path):
        total = 0
        sample_total = 0
        for x in range(0, pages):
            z = 0
            y = 0
            while True:
                if y == 3:
                    break
                elif z == 3:
                    print('!!Page Error!!')
                    break
                print('Page: ' + str(x + 1), '/', str(pages))
                i = 0
                y = 0
                while True:
                    if x == 0:
                        tree = grab(url + Date.path(date_on))
                        print(url + Date.path(date_on))
                    else:
                        tree = grab(url + Date.path(date_on) + path + str(x))
                        print(url + Date.path(date_on) + path + str(x))
                    listings = tree.xpath('//div[@itemprop="itemListElement"]')
                    if len(listings) == 0:
                        i += 1
                        print('--retrying--', str(i))
                        if i == 5:
                            print('--Page Error--')
                            break
                    else:
                        if not len(listings) == 0:
                            page_counter = ListingOnline.page_processor(listings, date_on)
                            if page_counter == int(-5):
                                print('--retrying--')
                                z += 1
                                break
                            else:
                                sample_total += len(listings)
                                total += page_counter
                                print(str(page_counter) + "/" + str(len(listings)))
                                y = 3
                                break
        return [total, sample_total]

    @staticmethod
    def scrape(url, path, period):
        length = len(period)
        for day in range(0, length):
            date_on = period[day]
            print('Date: ' + date_on.month + '/' + date_on.string_day)
            pages = get_pages(url + Date.path(date_on))
            if pages >= 18:
                print('Scrape- ALERT: Listings missed - Search span included more than 18 pages')
            elif pages == 0:
                pages += 1
            totals = ListingOnline.online_pages_iterate(date_on, pages, path)
            total = totals[0]
            sample_total = totals[1]
            if sample_total != 0:
                print(
                    'Sample of ' + str(total) + '/' + str(sample_total) + ' listings ('
                    + format(100 * total / sample_total, '.2f') + '% accuracy)\n')



#A few methods for ease of writing later code
def big_int(num):
    try:
        out = int(num)
    except ValueError:
        if "," in num:
            pre_out = num.replace(",", "")
            return int(pre_out)
    else:
        return out


def or_comma(item):
    top = len(item)
    count = 0
    for i in item:
        if i.isdigit() or i == ',':
            count += 1
    if count == top:
        return True
    else:
        return False


def month_length(date):
    month = date.month
    length = 0
    if int(month) == 1 or int(month) == 3 or int(month) == 5 or int(month) == 7 or int(
            month) == 8 or int(month) == 10 or int(month) == 12:
        length = 31
    elif int(month) == 4 or int(month) == 6 or int(month) == 9 or int(month) == 11:
        length = 30
    elif int(month) == 2:
        length = 28
    return length


def add_zero(num):
    if num < 10:
        return '0'+str(num)
    else:
        return str(num)


def if_zero(item):
    if item == 0:
        return True


def looking_forward(start, end):
    scan = []
    date_on = start
    while True:
        if not date_on.out() == end.out():
            scan.append(date_on)
            date_on = date_on.add_days(1)
            Date.scan.append(date_on)
        else:
            break
    if date_on.out() == end.out():
        scan.append(date_on)
        Date.scan.append(date_on)
    return scan


def days_forward(start, num):
    end = start.add_days(num)
    scan = looking_forward(start, end)
    return scan


def fileid_from_url(ob):
    """Return fileId from a url."""
    url = str(ob)
    index = url.find('~')
    fileId = url[index + 1:]
    local_id_index = fileId.find('/')

    share_key_index = fileId.find('?share_key')
    if share_key_index == -1:
        return fileId.replace('/', ':')
    else:
        return fileId[:share_key_index].replace('/', ':')


def my_round(x, base=5):
    return int(base * round(float(x) / base))

#HTML request
def grab(url):
    page = requests.get(url)
    soup = html.fromstring(page.content)
    return soup


#Returns the number of pages to iterate over for a given query
def get_pages(url):
    i = 0
    while True:
        page = requests.get(url)
        tree = html.fromstring(page.content)
        list_pages = tree.xpath('//li[@class = "buttonContainer_1am0dt"]/descendant::text()')
        try:
            string_pages = list_pages[-1]
        except:
            i += 1
            if i == 3:
                return 0
            continue
        else:
            pages = int(string_pages)
            return pages


def list_builder(range_start, range_length, item):
    output_list = []
    for i in range(range_start, range_length):
        if item == 'i':
            output_list.append(i)
        else:
            output_list.append(item)
    return output_list


#The run method generates the baseline database and the baseline distribution graph
def run(loc_title, path, searches):
    info = Listing.baseline(path, searches)
    total = info[0]
    sample_total = info[1]
    baseline_file_name = loc_title + ' Baseline'
    highest_bed_number = Listing.highest_bed_number
    highest_price = Listing.highest_price
    data = []
    global scheme
    for bed in range(1, highest_bed_number+1):
        out = Listing.list_by_bed(bed)
        name = str(bed) + ' Bed'
        if len(out) > 0:
            data.append(go.Histogram(x=out, marker=dict(color=scheme[bed]), opacity=0.65,
                                     name=name, xbins=dict(start=0, end=highest_price, size=20)))
    layout = go.Layout(barmode='overlay', xaxis=dict(tickprefix='$', fixedrange=True), yaxis=dict(fixedrange=True),
                       title=loc_title+' - Baseline - Sample of '+str(total)+'/'+str(sample_total)+' Listings')
    fig = go.Figure(data=data, layout=layout)
    return py.plot(fig, filename=baseline_file_name)


#Vacancy method genrates 30-day vacancy database and 30-day vacancy chart
def vacancy(loc_title, url, path, period):
    global scheme
    highest_bed_number = Listing.highest_bed_number
    month_listings = list_builder(0, len(period), 0)
    ListingOnline.scrape(url, path, period)
    graph_data = list()
    bed_number_count = 0
    xvals = list()
    for item in period:
        xvals.append(item.out())
    for bed_number in range(1, highest_bed_number + 1):
        if not Listing.count_by_bed(bed_number) > 3:
            continue
        else:
            month_data = []
            for day in range(len(period)):
                count = ListingOnline.count_by_bed_by_date(bed_number, period[day])
                month_data.append(count)
                month_listings[day] += count
            if not month_data == len(month_data)*[0]:
                bed_number_count += 1
                graph_data.append(go.Bar(x=xvals, y=month_data, marker=dict(color=scheme[bed_number-1]),
                                         name=str(bed_number) + ' Bed'))
            else:
                continue
    graph_data.append(go.Scatter(x=xvals, y=month_listings, name='Total Vacancy', line=dict(color=scheme, width=4)))

    layout = go.Layout(barmode='group', yaxis=dict(fixedrange=True), title=loc_title+' Vacancy Report', xaxis=dict(
        rangeslider=dict(),
        type='date', fixedrange=True
    ))

    fig = go.Figure(data=graph_data, layout=layout)
    name = loc_title + "Vacancy Report"

    return [py.plot(fig, filename=name), argrelextrema(np.asanyarray(month_listings), np.less)]


#Accesses ListingOnline database and plots distribution of price by bed-count
def run_a_date(loc_title, date_on):
    highest_bed_number = Listing.highest_bed_number
    group_labels = []
    hist_data = []
    for bed in range(1, highest_bed_number+1):
        out = ListingOnline.list_by_bed_by_date(bed, date_on)
        name = str(bed) + ' Bed'
        if len(out) > 3:
            hist_data.append(out)
            group_labels.append(name)

    fig = ff.create_distplot(hist_data, group_labels, show_hist=False, colors=scheme[:len(hist_data)])
    fig['layout'].update(title=date_on.out()+' - '+loc_title)
    name = loc_title + date_on.out()
    return py.plot(fig, filename=name)


#Initializes excel workbook for data output
def initialize_workbook(title):
    book = load_workbook(title+" Data.xlsx")
    writer = pd.ExcelWriter(title+" Data.xlsx", engine='openpyxl')
    writer.book = book
    writer.sheets = dict((ws.title, ws) for ws in book.worksheets)
    return writer


#Converts basline database to Pandas Dataframe
def baseline_frame(scan_length):
    baseline_data = ['ID', 'Title', 'Type', 'City', 'Price', 'Beds', 'Rating', 'Number of Reviews', '# of Vacant Days',
                     '% Vacancy']
    blank_entry = pd.Series(["--", "--", "--", "--", "--", "--", "--", "--", "--"])
    baseline_df = pd.DataFrame(data=[blank_entry, blank_entry], columns=baseline_data[1:], index=['blank', 'blank'])
    for item in Listing._baseline:
        sub = ListingOnline.search_by_id(item.id)
        if sub is None:
            num_of = "--"
            perc_of = "--"
        else:
            num_of = len(sub.vacant_dates.keys())
            perc_of = "{0:.1f}%".format(100*len(sub.vacant_dates.keys())/scan_length)

        temp = pd.DataFrame(data=[[item.title, item.typo, item.city, item.price, item.beds, item.rating,
                                   item.review_count, num_of, perc_of]], columns=baseline_data[1:], index=[item.id])
        baseline_df = baseline_df.append(temp)
    return baseline_df

#Converts vacancy database to Pandas Dataframe
def vacancy_frame(scan):
    cols = list()
    cols.append('')
    for item in scan:
        cols.append(item.out())
    blank_entry = pd.Series(len(scan)*["--"])
    vac_df = pd.DataFrame(data=[blank_entry, blank_entry], columns=cols[1:], index=['blank', 'blank'])
    for listing in ListingOnline.all_pulled:
        listing_data = []
        for i in range(1, len(cols)):
            unit = '--'
            for x in listing.vacant_dates:
                if x.out() == cols[i]:

                    unit = 'V'
            listing_data.append(unit)
        temp = pd.DataFrame(data=[listing_data], columns=cols[1:], index=[listing.id])
        vac_df = vac_df.append(temp)
    return vac_df


#Combines all prior methods into one method
#Takes location title, coordinates, base url, url extention, window size, name of excel document
#for output
def full_run(loc_title, coordinates, url, path, scan_length, book, new):
    start = Date.to_form(datetime.date.today() + datetime.timedelta(1))
    search_list = MapSearch.chop(MapSearch.clean(url, coordinates), 3)
    scan = days_forward(start, scan_length)
    plots = list()
    plots.append(run(loc_title, path, search_list))
    vac_results = vacancy(loc_title, url, path, scan)
    vac = vac_results[0]
    plots.insert(0, vac)
    minimum = vac_results[1][0]
    mins = []
    for num in minimum:
        mins.append(Date.scan[num-1])
    for item in mins:
        plots.append(run_a_date(loc_title, item))
    out = []
    for item in plots:
        out.append(fileid_from_url(item))
    end_time = datetime.datetime.now()
    global start_time
    diff = end_time - start_time
    print("Time Elapsed: "+str(diff))
    if not new == 'new':
        book = initialize_workbook(loc_title)
    baseline_frame(scan_length).to_excel(book, start.out()+' Baseline')

    return out


location_title = "Boca Raton"
location = "boca-raton"
coords = 'ne_lat=18.478222609602486&ne_lng=-66.10346080373603&sw_lat=18.456396421667705&sw_lng=-66.12285853933173'
url = "https://www.airbnb.com/s/"+location+"?room_types%5B%5D=Entire%20home%2Fapt"
path = "&section_offset="
workbook = pd.ExcelWriter(location_title+" Data.xlsx")
