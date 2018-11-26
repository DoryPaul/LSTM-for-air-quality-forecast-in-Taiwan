from cmath import sqrt
import  numpy as np
from numpy import concatenate
from matplotlib import pyplot
from pandas import read_csv
from pandas import DataFrame
from pandas import concat
from sklearn.preprocessing import MinMaxScaler
from sklearn.preprocessing import LabelEncoder
from sklearn.metrics import mean_squared_error
from keras.models import Sequential
from keras.layers import Dense
from keras.layers import LSTM
import xlrd
import xlwt




# convert series to supervised learning
def series_to_supervised(data, n_in=1, n_out=1, dropnan=True):
    n_vars = 1 if type(data) is list else data.shape[1]

    df = DataFrame(data)
    cols, names = list(), list()
    # input sequence (t-n, ... t-1)
    for i in range(n_in, 0, -1):

        cols.append(df.shift(i))
        names += [('var%d(t-%d)' % (j + 1, i)) for j in range(n_vars)]
    # forecast sequence (t, t+1, ... t+n)
    for i in range(0, n_out):
        cols.append(df.shift(-i))
        if i == 0:
            names += [('var%d(t)' % (j + 1)) for j in range(n_vars)]
        else:
            names += [('var%d(t+%d)' % (j + 1, i)) for j in range(n_vars)]
    # put it all together
    agg = concat(cols, axis=1)

    agg.columns = names

    # drop rows with NaN values
    if dropnan:
        agg.dropna(inplace=True)
    return agg




def data_deal(values):
    # integer encode direction
    #encoder = LabelEncoder()
    #values[:, 0] = encoder.fit_transform(values[:, 0])

# ensure all data is float
    values = values.astype('float32')
# normalize features
    scaler = MinMaxScaler(feature_range=(0, 1))
    scaled = scaler.fit_transform(values)

# frame as supervised learning
    reframed = series_to_supervised(scaled, 1, 1)
# drop columns we don't want to predict
    reframed.drop(reframed.columns[[12, 13, 14, 15,16,17,18,19,20,21]], axis=1, inplace=True)
    return reframed,scaler

def two_in_one(values1,model):
    temp = values1
    reframed1, scaler1 = data_deal(values1)
    values1 = reframed1.values
    values1_X = values1[:, :-1]
    values1_X = values1_X.reshape((values1_X.shape[0], 1, values1_X.shape[1]))
    yhat = model.predict(values1_X)
    # invert scaling for actual
    values1_X = values1_X.reshape((values1_X.shape[0], values1_X.shape[2]))
    # invert scaling for forecast
    inv_yhat = concatenate((yhat, values1_X[:, 1:]), axis=1)
    inv_yhat = scaler1.inverse_transform(inv_yhat)
    #inv_yhat = inv_yhat[:, 0]
    inv_yhat=np.vstack((temp[-1],inv_yhat))
    return inv_yhat
# load dataset
dataset = read_csv('E:\shaluxz.csv', header=0, index_col=0)
values = dataset.values

sampleno = 365*24*3-1
mu = 0
sigma = 1
np.random.seed(0)
s = np.random.normal(mu,sigma,sampleno)
c = values[365*24*2+1:]
normaldis=[]
for i in range(len(s)):
    c[i][1] += s[i]

reframed,scaler=data_deal(values)

print(reframed.head())

# split into train and test sets
values = reframed.values

n_train_hours = 365 * 24 * 2
train = values[:n_train_hours, :]
test = values[n_train_hours:, :]

# split into input and outputs
train_X, train_y = train[:, :-1], train[:, -1]
test_X, test_y = test[:, :-1], test[:, -1]
# reshape input to be 3D [samples, timesteps, features]
train_X = train_X.reshape((train_X.shape[0], 1, train_X.shape[1]))
test_X = test_X.reshape((test_X.shape[0], 1, test_X.shape[1]))

#testcase with normaldistribution
reframedc,scalerc=data_deal(c)
c = reframed.values
c_X, c_y = c[:, :-1], c[:, -1]
c_X = c_X.reshape((c_X.shape[0], 1, c_X.shape[1]))


# design network
model = Sequential()
model.add(LSTM(50, input_shape=(train_X.shape[1], train_X.shape[2])))
model.add(Dense(1))
model.compile(loss='mae', optimizer='adam')
# fit network
history = model.fit(train_X, train_y, epochs=1, batch_size=100, validation_data=(test_X, test_y), verbose=2,
                    shuffle=False)

# evaluate network
evaluation = model.evaluate(train_X,train_y,batch_size=32, verbose=2, sample_weight=None)
print('The evaluation is ',evaluation)
# plot history
pyplot.plot(history.history['loss'], label='train')
pyplot.plot(history.history['val_loss'], label='test')
pyplot.legend()
pyplot.show()


# make a prediction
yhat = model.predict(test_X)

test_X = test_X.reshape((test_X.shape[0], test_X.shape[2]))
# invert scaling for forecast
inv_yhat = concatenate((yhat, test_X[:, 1:]), axis=1)
inv_yhat = scaler.inverse_transform(inv_yhat)
inv_yhat = inv_yhat[:, 0]


# invert scaling for actual
test_y = test_y.reshape((len(test_y), 1))
inv_y = concatenate((test_y, test_X[:, 1:]), axis=1)
inv_y = scaler.inverse_transform(inv_y)
inv_y = inv_y[:, 0]

# make a prediction of c
chat = model.predict(c_X)
c_X = c_X.reshape((c_X.shape[0], c_X.shape[2]))
# invert scaling for forecast
inv_chat = concatenate((chat, c_X[:, 1:]), axis=1)
inv_chat = scalerc.inverse_transform(inv_chat)
inv_chat = inv_chat[:, 0]
rmsec = sqrt(mean_squared_error(c[0:,1],inv_chat))
print('RMSEC: %.3f' % rmsec.real)
# calculate RMSE
rmse = sqrt(mean_squared_error(inv_y, inv_yhat))

print('Test RMSE: %.3f' % rmse.real)
file = xlwt.Workbook()
table = file.add_sheet('宜兰1.xls')
for i in range(len(inv_y)):
    table.write(i,0,float(inv_y[i]))
    table.write(i,1,float(inv_yhat[i]))
file.save(r'E:\宜兰1.xls')


#print('测试1',values1)
#testdata=two_in_one(values1,model)
#print('测试',testdata.shape)
#test2=two_in_one(testdata,model)
#print('测试2',test2.shape)
def rec(value,n,model,result):
    a=two_in_one(value,model)
    result.extend(list(a[1:,0]))
    #print('第{0}次预测结果：{1}'.format(n+1,a[1:,0]))
    n+=1
    if n<24:
        return rec(a,n,model,result)
    else:
        return result

x1 = []
x2 = dataset.values[35040:,1]

for i in range(365):
    x=35038+i*24
    y = 35040 + i * 24
    values1=dataset.values[x:y]
    b=rec(values1,0,model,[])
    x1.extend(b)
print('x1',x1)

pyplot.plot(x1, label='predict')
pyplot.plot(x2, label='real')
pyplot.legend()
pyplot.show()
