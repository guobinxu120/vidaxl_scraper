# -*- coding: utf-8 -*-
from scrapy import Spider, Request
from collections import OrderedDict
from xlrd import open_workbook
import os, requests, re, time, json
from scrapy.http import TextResponse

def download(url, destfilename):
    if not os.path.exists(destfilename):

        try:
            r = requests.get(url, stream=True)
            # if r.status_code != 200:
            #     r = requests.get(temp_url, stream=True)
            with open(destfilename, 'wb') as f:
                for chunk in r.iter_content(chunk_size=1024):
                    if chunk:
                        f.write(chunk)
                        f.flush()
        except:
            # print(url)
            pass

def readExcel(path):
    wb = open_workbook(path)
    result = []
    for sheet in wb.sheets():
        number_of_rows = sheet.nrows
        number_of_columns = sheet.ncols
        herders = []
        for row in range(0, number_of_rows):
            values = OrderedDict()
            for col in range(number_of_columns):
                value = (sheet.cell(row,col).value)
                if row == 0:
                    herders.append(value)
                else:

                    values[herders[col]] = value
            if len(values.values()) > 0:
                result.append(values)
        break

    return result


class AngelSpider(Spider):
    name = "vidaxl_scraper"
    start_urls = 'https://www.vidaxl.fr/'
    count = 0
    use_selenium = False
    urls = readExcel("Input.xlsx")
    models = []

    headers = ['EAN', 'SKU', 'URL', 'Nom', 'Prix', 'Categorie1','Categorie2', 'Categorie3', 'Categorie4', 'Categorie5','Categorie6',
               'Vendu par', 'Délai de livraison', 'Description', 'Image1', 'Image2', 'Image3', 'Image4', 'Image5']

    def start_requests(self):
        yield Request(self.start_urls, callback=self.parse)

    def parse(self, response):
        for i, val in enumerate(self.urls):
            ern_code = val['URL']
            if ern_code != '':
                yield Request(response.urljoin(ern_code), callback=self.parse1)
                # yield Request('https://www.manomano.fr/manuel-et-livre-de-jardinage-3811', callback=self.parse1)
    def parse1(self, response):
        try:
            body_json = json.loads(response.body)
            body_text = body_json['page']
            response = TextResponse(url='',
                                body= body_text,
                                encoding='utf-8')
        except:
            pass
        urls = response.xpath('//div[@class="items"]//div[@class="grid-view"]/a/@href').extract()
        for url in urls:
            time.sleep(1)
            # pass
            yield Request(response.urljoin(url), callback=self.final_parse)
                # yield Request('https://www.manomano.fr/jardinage-pour-enfants-3867', callback=self.final_parse)
                # break
        next_tag = response.xpath('//div[@id="show-more-product"]/span/@data').extract_first()
        if next_tag:
            yield Request(next_tag + '&is_ajax=1', callback=self.parse1)

    def final_parse(self, response):
        # try:
            item = OrderedDict()
            for key in self.headers:
                item[key] = None

            categories = response.xpath('//div[@id="breadcrumbs"]/ul/li/a/span/text()').extract()
            for i, category in enumerate(categories):
                if i == len(categories) -1 : break
                item['Categorie{}'.format(str(i+1))] = category

            name = response.xpath('//div[@class="container-top"]/h1/text()').extract_first().strip()
            # name = ' '.join(name_list).replace('\r\n', '').replace('  ', ' ').strip()
            item['Nom'] = name
            item['URL'] = response.url

            price = response.xpath('//div[@class="price-show"]/meta[@itemprop="price"]/@content').extract_first()
            item['Prix'] = price

            sku = response.xpath('//div[@class="price-show"]/meta[@itemprop="sku"]/@content').extract_first()
            item['SKU'] = sku
            ean = response.xpath('//div[@class="price-show"]/meta[@itemprop="gtin13"]/@content').extract_first()
            item['EAN'] = ean

            item['Vendu par'] = response.xpath('//span[@class="seller-name"]/text()').extract_first().strip()
            item['Délai de livraison'] = response.xpath('//div[@class="delivery-name"]/text()').extract_first().replace('Délai de livraison', '').replace(':', '').strip()

            descript1_list = response.xpath('//div[@class="product-description-content"]//text()').extract()
            descript1 = ''.join(descript1_list).strip()
            item['Description'] = descript1

            tr_tags = response.xpath('//ul[@class="specs"]/li/text()').extract()
            for tr in tr_tags:
                keys = tr.split(':')
                if len(keys) > 1:
                    key = keys[0].strip()
                    val = keys[1].strip()
                    item[key] = val
                if not key in self.headers:
                    self.headers.append(key)


            for i in range(5):
                item['Image' + str(i+1)] = ''

            index = 0
            image_urls = response.xpath('//ul[contains(@class, "swiper-wrapper")]/li/img/@src').extract()

            for image_url in image_urls:
                index += 1
                if index > 5:
                    break
                temp_name = re.sub('[^A-Za-z0-9 ]+', '', name)
                image_name = 'tropicmarket_{}_{}_{}.jpg'.format(str(index), temp_name, item['EAN'])
                download(image_url.replace('_thumb', '_hd').strip(), 'Images/'+image_name)
                item['Image' + str(index)] = 'http://tic-et-tec.com/TTM/' + image_name

            self.models.append(item)
            self.count += 1
            print(self.count)
            yield item

        # except :
        #     pass

    def getImage(self, response):
        item = response.meta['item']
        index = 0
        img_urls = response.meta['img_urls']
        image_urls = response.xpath('//div[@id="darty_zoom_popin_container"]//img/@src').extract()

        for i, image_url in enumerate(image_urls):
            index += 1
            if index > 5:
                break
            temp_name = re.sub('[^A-Za-z0-9 ]+', '', item['Nom'])
            image_name = temp_name.strip().replace(' ', '_').replace('__', '_') + '_' + 'tropicmarket' + '_' + str(index) +".jpg"
            filename = "Images/" + image_name
            tem_url = img_urls[i]
            download(image_url.strip(), filename, tem_url)
            item['Image' + str(index)] = image_name

        self.models.append(item)
        yield item