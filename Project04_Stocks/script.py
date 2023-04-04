import yfinance as yf

def loading_data(ticker):
    df = yf.Ticker(ticker).history("1y")
    df.reset_index(inplace=True)
    df['ticker'] = ticker.split(".")[0]
    return df

petrobras = loading_data("PETR4.SA")
banco_brasil = loading_data("BBAS3.SA")
itau = loading_data("ITUB")
copel = loading_data("CPLE3.SA")
sanepar = loading_data("SAPR4.SA")
vale = loading_data("VALE3.SA")
klabin = loading_data("KLBN4.SA")
taesa = loading_data("TAEE4.SA")

