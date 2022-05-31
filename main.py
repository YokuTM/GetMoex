import requests
import apimoex
import pandas as pd
import plotly.graph_objects as go


pd.set_option('display.max_rows', None)
pd.set_option('display.max_columns', None)
pd.set_option('display.max_colwidth', None)

moexname = input("Введите название акции:")
startdate = input("Введите дату начала в формате ГГГГ-ММ-ДД:")
enddate = input("Введите дату конца в формате ГГГГ-ММ-ДД:")

with requests.Session() as session:
        data = apimoex.get_board_history(session, moexname, start=startdate, end=enddate,
                                         columns=('TRADEDATE', 'OPEN', 'HIGH', 'LOW', 'CLOSE', 'VOLUME'))
        df = pd.DataFrame(data)
        df.set_index('TRADEDATE', inplace=True)
        # print(df)
        df.to_excel('test.xlsx', sheet_name=moexname, index=True)

    # Постройка свечей и МА
df = pd.read_excel('test.xlsx')
df['MA'] = df['CLOSE'].rolling(5).mean()
fig = go.Figure(data=[go.Candlestick(x=df['TRADEDATE'],
                                         open=df['OPEN'],
                                         high=df['HIGH'],
                                         low=df['LOW'],
                                         close=df['CLOSE'])])
fig.add_trace(
        go.Scatter(
            x=df['TRADEDATE'],
            y=df['MA'],
            line=dict(color="#0000ff"),
            name="MA"
        ))

fig.write_html('graph.html')



