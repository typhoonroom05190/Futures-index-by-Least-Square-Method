def Least_square_method(data):
    if len(data) == 1:
        return int(data[0])
    if len(data) > 6:
        data = data[-7:-1]
    x = np.arange(0,len(data))
    y = np.array(data)
    N = len(y)
    B = (sum(x[i] * y[i] for i in x) - 1./N*sum(x)*sum(y)) / (sum(x[i]**2 for i in x) - 1./N*sum(x)**2)
    A = 1.*sum(y)/N - B * 1.*sum(x)/N
    return int(A + B * N)
