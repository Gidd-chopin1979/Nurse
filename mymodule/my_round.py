#'AngleX(deg)'において負の値を四捨五入するために，新しいround関数を定義
# 引数をfloatにして，小数点以下が２桁残るように四捨五入
def main(value):
    value = float(value)
    if value <= 0:
        value = round(abs(value),2)
        value = value*(-1)
    else:
        value = round((value),2)
    return(value)

if __name__ == '__name__':
    main()
