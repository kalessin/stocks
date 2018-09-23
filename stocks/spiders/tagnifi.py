from scrapy import Spider, Request
from scrapy.exceptions import NotConfigured


class TagnifiSpider(Spider):
    
    name = "tagnifi"
    BASE_URL = ('https://viewer.tagnifi.com/api/fundamentals?company={company}&statement={statement}'
               '&period_type={period}&relative_period=0&limit={limit}&industry_template=commercial')
    STATEMENTS = [('balance_sheet_statement', 'quarter'),
                  ('income_statement', 'ttm'),
                ('cash_flow_statement', 'ttm')]

    
    company = None
    limit = 1

    custom_settings = {
        "CONCURRENT_REQUEST": 1,
        "DOWNLOAD_DELAY": 20,
    }

    def start_requests(self):
        if self.company is None:
            raise NotConfigured
        for statement, period in self.STATEMENTS:
            yield Request(url=self.BASE_URL.format(company=self.company, statement=statement, period=period,
                                                   limit=self.limit),
                          meta={'statement': statement})
    
    def parse(self, response):
        statement = response.meta['statement']
        open(f"{self.company}-{statement}.json", "w").write(response.text)
