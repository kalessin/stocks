from scrapy import Spider, Request
from scrapy.exceptions import NotConfigured


AVL_PERIODS = {'quarter', 'ttm', 'annual'}


class TagnifiSpider(Spider):
    
    name = "tagnifi"
    BASE_URL = ('https://viewer.tagnifi.com/api/fundamentals?company={company}&statement={statement}'
               '&period_type={period_type}&relative_period=0&limit={limit}&industry_template=commercial')
    STATEMENTS = [('balance_sheet_statement', 'quarter'),
                  ('income_statement', 'ttm'),
                ('cash_flow_statement', 'ttm')]

    
    companies = None  # comma separated list of companies ticks
    limit = 1
    period_type = None

    # be friendly!
    custom_settings = {
        "CONCURRENT_REQUEST": 1,
        "DOWNLOAD_DELAY": 20,
    }

    def start_requests(self):
        if not self.companies:
            raise NotConfigured
        for statement, period_type in self.STATEMENTS:
            period_type = self.period_type or period_type
            for company in self.companies.split(','):
                yield Request(url=self.BASE_URL.format(company=company, statement=statement, period_type=period_type,
                                                       limit=self.limit),
                              meta={'statement': statement, 'period_type': period_type, 'limit': self.limit,
                                    'company': company})
    
    def parse(self, response):
        statement = response.meta['statement']
        period_type = response.meta['period_type']
        limit = response.meta['limit']
        company = response.meta['company']
        open(f"{company}-{statement}-{period_type}-{limit}.json", "w").write(response.text)
