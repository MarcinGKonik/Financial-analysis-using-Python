import matplotlib.pyplot as plt

#Visualizing stock data
def ChartGen(ticker, hist):

    plt.figure(figsize=(10, 5))

    plt.plot(
        hist.index,
        hist['Close'],
        label='Close Price',
        linewidth=2
    )

    plt.title(f"{ticker} - graph")
    plt.xlabel("Date")
    plt.ylabel("Price")
    plt.legend()
    plt.grid(True)

    print("Visual data plot generated")

    chart_path = "price_chart.png"

    plt.savefig(chart_path)
    plt.close()

    print("Visual data plot saved")

    return chart_path