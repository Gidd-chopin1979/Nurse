#hh:mm:ss表記の時刻を，:無し，単位sの表記に変換する
#'AngleX(deg)'において負の値を四捨五入するために，新しいround関数を定義
# 引数をfloatにして，小数点以下が２桁残るように四捨五入
def main(value):
    hh, mm, ss = map(float, value.split(':'))
    return ss + 60*(mm + 60*hh)

if __name__ == '__name__':
    main()