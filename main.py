import requests
import apimoex
import pandas as pd
import plotly.graph_objects as go
from tkinter import *
from tkinter import messagebox

# Настройка формы
root = Tk()


def btn_click():
    pd.set_option('display.max_rows', None)
    pd.set_option('display.max_columns', None)
    pd.set_option('display.max_colwidth', None)

    moexname = mName.get()
    startdate = sDate.get()
    enddate = eDate.get()

    with requests.Session() as session:
        data = apimoex.get_board_history(session, moexname, start=startdate, end=enddate,
                                         columns=('TRADEDATE', 'OPEN', 'HIGH', 'LOW', 'CLOSE', 'VOLUME'))
        df = pd.DataFrame(data)
        df.set_index('TRADEDATE', inplace=True)
        # print(df)
        df.to_excel('test.xlsx', sheet_name=moexname, index=True)

    info_str = f'Данные акций "{str(moexname)}" за период с {str(startdate)} по {str(enddate)} сохранены.'
    messagebox.showinfo(title='Info', message=info_str)

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


# Настройка интерфейса
root['bg'] = '#fafafa'
root.title('pMoex')
root.wm_attributes('-alpha', 0.9)
root.geometry('400x230')
root.resizable(width=False, height=False)

canvas = Canvas(root, height=400, width=230)
canvas.pack()

frame = Frame(root)
frame.place(relwidth=1, relheight=1)

title = Label(frame, text='Введите название акции:', font=40)
title.place(x=20, y=10, height=20, width=360)
mName = Entry(frame, bg='white')
mName.place(x=20, y=40, height=20, width=360)
title1 = Label(frame, text='Введите дату начала в формате ГГГГ-ММ-ДД:', font=40)
title1.place(x=20, y=70, height=20, width=360)
sDate = Entry(frame, bg='white')
sDate.place(x=20, y=100, height=20, width=360)
title2 = Label(frame, text='Введите дату окончания в формате ГГГГ-ММ-ДД:', font=40)
title2.place(x=20, y=130, height=20, width=360)
eDate = Entry(frame)
eDate.place(x=20, y=160, height=20, width=360)
btn = Button(frame, text='Старт', command=btn_click)
btn.place(x=20, y=190, height=30, width=360)

root.mainloop()
