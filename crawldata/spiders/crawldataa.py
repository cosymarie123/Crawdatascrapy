import scrapy



class CrawldataaSpider(scrapy.Spider):
    name = 'crawldataa'
    page_number = 2
    def start_requests(self):   
        start_urls = [
            'https://homedy.com/ban-can-ho-chung-cu/p1',
        ]
        for url in start_urls:
            yield scrapy.Request(url=url, callback=self.parse)  
    def parse(self, response):
        for quote in response.css(".product-item-top"):
            yield {
                'link': quote.css('a::attr(href)')[0].extract()
            }

        next_page = 'https://homedy.com/ban-can-ho-chung-cu/p' + str(CrawldataaSpider.page_number)
        if CrawldataaSpider.page_number <= 10:
            CrawldataaSpider.page_number +=1           
            yield response.follow(next_page, callback = self.parse)      

    pass
