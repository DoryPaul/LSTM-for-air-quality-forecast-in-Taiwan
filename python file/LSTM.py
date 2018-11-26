from cmath import sqrt
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
    encoder = LabelEncoder()
    values[:, 0] = encoder.fit_transform(values[:, 0])

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

# load dataset
dataset = read_csv('E:\shaluxz.csv', header=0, index_col=0)
values = dataset.values
values1=dataset.values[43797:43799]
reframed,scaler=data_deal(values)
reframed1,scaler1=data_deal(values1)

print('test',reframed1.values)
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

values1=reframed1.values
values1_X=values1[:,:-1]
values1_X = values1_X.reshape((values1_X.shape[0], 1, values1_X.shape[1]))
#print(train_X.shape, train_y.shape, test_X.shape, test_y.shape)

# design network
model = Sequential()
model.add(LSTM(50, input_shape=(train_X.shape[1], train_X.shape[2])))
model.add(Dense(1))
model.compile(loss='mae', optimizer='adam')
# fit network
history = model.fit(train_X, train_y, epochs=1, batch_size=100, validation_data=(test_X, test_y), verbose=2,
                    shuffle=False)

# evaluate network
evaluation=model.evaluate(train_X,train_y,batch_size=32, verbose=2, sample_weight=None)
print('The evaluation is ',evaluation)
# plot history
pyplot.plot(history.history['loss'], label='train')
pyplot.plot(history.history['val_loss'], label='test')
pyplot.legend()
pyplot.show()


# make a prediction
yhat = model.predict(test_X)
b=model.predict(values1_X)
print(yhat)
print('Test',b)

test_X = test_X.reshape((test_X.shape[0], test_X.shape[2]))
# invert scaling for forecast
inv_yhat = concatenate((yhat, test_X[:, 1:]), axis=1)
inv_yhat = scaler.inverse_transform(inv_yhat)
inv_yhat = inv_yhat[:, 0]

values1_X =values1_X.reshape((values1_X.shape[0],values1_X.shape[2]))
ib=concatenate((b,values1_X[:,1:]),axis=1)
ib=scaler1.inverse_transform(ib)
ib=ib[:,0]
print('Testib',ib)

# invert scaling for actual
test_y = test_y.reshape((len(test_y), 1))
inv_y = concatenate((test_y, test_X[:, 1:]), axis=1)
inv_y = scaler.inverse_transform(inv_y)
inv_y = inv_y[:, 0]

# calculate RMSE
rmse = sqrt(mean_squared_error(inv_y, inv_yhat))

print('Test RMSE: %.3f' % rmse.real)
file = xlwt.Workbook()
table = file.add_sheet('宜兰1.xls')
for i in range(len(inv_y)):
    table.write(i,0,float(inv_y[i]))
    table.write(i,1,float(inv_yhat[i]))
file.save(r'E:\宜兰1.xls')
