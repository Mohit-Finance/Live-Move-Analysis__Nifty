import os
import sys
import time
import math
import json
import threading
from datetime import datetime, date, timedelta

import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.gridspec as gridspec
from scipy.stats import norm

import requests
import pyotp
import xlwings as xw
import upstox_client

import undetected_chromedriver as uc
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

lock = threading.Lock()
india_vix_ready = threading.Event()

live_data = {}
dict_lock = threading.Lock()
excel_lock = threading.Lock()
access = None

##############################################################################

tdate = datetime.now().date()
code = None
access = None

os.makedirs(f'Credentials/Data/{tdate}', exist_ok=True)

def time_fun():
	ttime = datetime.now().time().replace(microsecond=0)
	ttime = ttime.strftime("%H:%M:%S")
	return ttime

def show_totp(secret):
	totp = pyotp.TOTP(secret)
	otp = totp.now()
	return otp

if not os.path.exists('Credentials/login_details.json'):
	print("User Details not found. First Create a User Base & Retry. Exiting program.")
	sys.exit()

with open('Credentials/login_details.json', 'r') as file_read:
	users_data = json.load(file_read)

allowed_namess = users_data.keys()
allowed_names = [name.lower() for name in allowed_namess]

name_dict = {}

for i in range(len(allowed_names)):
	name_dict[f'{allowed_names[i]}'] = f'{tdate}_access_code_{allowed_names[i]}.json'

name_list = name_dict.values()

file_list = os.listdir(f'Credentials/Data/{tdate}')

for name in name_list:
	if name in file_list:
		with open(f'Credentials/Data/{tdate}/{name}', 'r') as file_read:
			access = json.load(file_read)
			acc_name = name[23:][:-5]

if not access:

	while True:
		acc_name = input(f'\nEnter Name of Account Holder to Login From {list(allowed_namess)} : ').lower()
		if acc_name in allowed_names:
			break
		else:
			print(f"\nInvalid User. Please Enter Registered User Name {list(allowed_namess)}'.")

	try:
		with open(f'Credentials/Data/{tdate}/{tdate}_access_code_{acc_name}.json', 'r') as file_read:
			access = json.load(file_read)

	except:

		with open('Credentials/login_details.json', 'r') as file_read:
			login_details = json.load(file_read)

		api_key = login_details[f'{acc_name.capitalize()}']['api_key']
		api_secret = login_details[f'{acc_name.capitalize()}']['api_secret']
		api_auth = login_details[f'{acc_name.capitalize()}']['api_auth']
		api_pin = login_details[f'{acc_name.capitalize()}']['pin']
		mobile_no = login_details[f'{acc_name.capitalize()}']['Mob No.']
		hold_name = login_details[f'{acc_name.capitalize()}']['full_name']

		print(f'\nTrying to Login from Account Holder: {hold_name}')

		uri = 'https://www.google.com/'
		url1 = f'https://api.upstox.com/v2/login/authorization/dialog?response_type=code&client_id={api_key}&redirect_uri={uri}\n'

		options = uc.ChromeOptions()
		options.headless = True
		options.add_argument("--disable-gpu")
		options.add_argument("--no-sandbox")
		driver = uc.Chrome(options=options)

		# driver = uc.Chrome() # Use this line instead to run Chrome in normal (visible) mode, (In that case, comment out the 5 lines above that set headless options)

		driver.get(url1)
		wait = WebDriverWait(driver, 20)
		phone_input = wait.until(EC.presence_of_element_located((By.ID, "mobileNum")))
		phone_input.send_keys(mobile_no)
		otp_button = wait.until(EC.element_to_be_clickable((By.ID, "getOtp")))
		otp_button.click()
		# print("✅ Phone number entered, now captcha should appear normally")

		totp_value = show_totp(api_auth)
		totp_input = wait.until(EC.presence_of_element_located((By.ID, "otpNum")))
		totp_input.send_keys(totp_value)
		proceed_button = wait.until(EC.element_to_be_clickable((By.ID, "continueBtn")))
		proceed_button.click()
		# print("✅ TOTP entered and Continue clicked!")

		pin_input = wait.until(EC.presence_of_element_located((By.ID, "pinCode")))
		pin_input.send_keys(api_pin)
		proceed_button = wait.until(EC.element_to_be_clickable((By.ID, "pinContinueBtn")))
		proceed_button.click()

		# print("✅ PIN entered and proceed button clicked!")
		time.sleep(3)
		code_url = driver.current_url

		driver.quit()

		start = code_url.find('code=')
		if start != -1:
			start =start + 5  # move past 'code='
			code = code_url[start:start+6]
		else:
			print("No code found in the URL")

		url = 'https://api.upstox.com/v2/login/authorization/token'
		headers = {
			'accept': 'application/json',
			'Content-Type': 'application/x-www-form-urlencoded',
		}

		data = {
			'code': code,
			'client_id': api_key,
			'client_secret': api_secret,
			'redirect_uri': uri,
			'grant_type': 'authorization_code',
		}

		response = requests.post(url, headers=headers, data=data)
		access = response.json()['access_token']
		print(f'\nLogin Successful, Status Code : {response.status_code}')
		print(f"User Name : {response.json()['user_name']}\nEmail ID : {response.json()['email']}")

		with open(f'Credentials/Data/{tdate}/{tdate}_access_code_{acc_name}.json', 'w') as file_write:
			json.dump(access, file_write)

print(f'\nLogin Successful from Account : {acc_name.capitalize()}')


######################################################################################

live_data = {'nifty':{'open': None, 'high':None, 'low':None, 'close':None}, 'india_vix':{'open': None, 'high':None, 'low':None, 'close':None}}
nifty_openn=nifty_high=nifty_low=nifty_close=india_vix_open=india_vix_high=india_vix_low=india_vix_close=None

def get_ohlc_by_interval(message, symbol, interval):
	feed = (
		message.get("feeds", {})
			   .get(symbol, {})
			   .get("fullFeed", {})
	)

	# Nifty uses indexFF, VIX does not
	if "indexFF" in feed:
		feed = feed.get("indexFF", {})

	ohlc_list = feed.get("marketOHLC", {}).get("ohlc", [])

	for candle in ohlc_list:
		if candle.get("interval") == interval:
			return candle

	return None


def on_message(message):
	global live_data
	global nifty_openn, nifty_high, nifty_low, nifty_close
	global india_vix_open, india_vix_high, india_vix_low, india_vix_close

	if message.get("type") != "live_feed":
		return

	if 'NSE_INDEX|Nifty 50' in message['feeds']:
		nifty_ohlc = get_ohlc_by_interval(message, "NSE_INDEX|Nifty 50", "1d")
		if nifty_ohlc is not None:
			nifty_openn = nifty_ohlc.get("open")
			nifty_high  = nifty_ohlc.get("high")
			nifty_low   = nifty_ohlc.get("low")
			nifty_close = nifty_ohlc.get("close")

	if 'NSE_INDEX|India VIX' in message['feeds']:
		daily_ohlc = get_ohlc_by_interval(message, "NSE_INDEX|India VIX", "1d")
		if daily_ohlc is not None:
			india_vix_open  = daily_ohlc.get("open")
			india_vix_high  = daily_ohlc.get("high")
			india_vix_low   = daily_ohlc.get("low")
			india_vix_close = daily_ohlc.get("close")

	with lock:
		if nifty_openn is not None:
			live_data['nifty']['open'] = nifty_openn
			live_data['nifty']['high'] = nifty_high
			live_data['nifty']['low'] = nifty_low
			live_data['nifty']['close'] = nifty_close

		if india_vix_open is not None:
			live_data['india_vix']['open'] = india_vix_open
			live_data['india_vix']['high'] = india_vix_high
			live_data['india_vix']['low'] = india_vix_low
			live_data['india_vix']['close'] = india_vix_close

streamer = None
def main():
	global streamer
	configuration = upstox_client.Configuration()
	access_token = access
	configuration.access_token = access_token

	streamer = upstox_client.MarketDataStreamerV3(
					upstox_client.ApiClient(configuration), ['NSE_INDEX|Nifty 50', 'NSE_INDEX|India VIX'], "full")

	streamer.on("message", on_message)

	streamer.connect()


if __name__ == "__main__":
	main()
#########################################################################################

while not live_data['india_vix']['close']:
	print('Waiting for Websocet Data Receive')
	time.sleep(1)

print('Websocket Data Receiving Started')

configuration = upstox_client.Configuration()
configuration.access_token = access

api = upstox_client.HistoryV3Api(upstox_client.ApiClient(configuration))

to_date= date.today()
from_date = to_date - timedelta(days=3000)

to_date   = to_date.strftime("%Y-%m-%d")
from_date = from_date.strftime("%Y-%m-%d")


try:
	response_daily = api.get_historical_candle_data1(instrument_key="NSE_INDEX|Nifty 50", unit="days", interval="1", from_date=from_date, to_date=to_date)
	response_daily = response_daily.data.candles
	df_daily = pd.DataFrame(response_daily, columns=["time", "open", "high", "low", "close", "vol1", "vol2"]).drop(['vol1', 'vol2'], axis='columns')

	response_weeks = api.get_historical_candle_data1(instrument_key="NSE_INDEX|Nifty 50", unit="weeks", interval="1", from_date=from_date, to_date=to_date)
	response_weeks = response_weeks.data.candles[1:]
	df_weeks = pd.DataFrame(response_weeks, columns=["time", "open", "high", "low", "close", "vol1", "vol2"]).drop(['vol1', 'vol2'], axis='columns')

	response_months = api.get_historical_candle_data1(instrument_key="NSE_INDEX|Nifty 50", unit="months", interval="1", from_date=from_date, to_date=to_date)
	response_months = response_months.data.candles[1:]
	df_months = pd.DataFrame(response_months, columns=["time", "open", "high", "low", "close", "vol1", "vol2"]).drop(['vol1', 'vol2'], axis='columns')

except Exception as e:
	print(f'Error : {e}')

df_dict = {'daily': {'percentile':[], 'Low_side':[], 'High_side':[], 'Points_UP': [], 'Points_Down': []},
		   'weekly': {'percentile':[], 'Low_side':[], 'High_side':[], 'Points_UP': [], 'Points_Down': []},
		   'monthly': {'percentile':[], 'Low_side':[], 'High_side':[], 'Points_UP': [], 'Points_Down': []}}

names = ['daily', 'weekly', 'monthly']

for j, data in enumerate([df_daily, df_weeks, df_months]):

	data['prev_close'] = data['close'].shift(-1)
	# data['prev_close'] = data['open']
	data['up_move'] = np.maximum((data['high'] - data['prev_close']),0)
	data['down_move'] = np.maximum((data['prev_close'] - data['low']),0)
	data['up_pct'] = data['up_move']/data['prev_close']
	data['down_pct'] = data['down_move']/data['prev_close']
	data['Total Range'] = data['up_move'] + data['down_move']
	data.dropna(inplace=True)

	up_data = data['up_pct'].tolist()
	down_data = data['down_pct'].tolist()
	up_points = data['up_move'].tolist()
	down_points = data['down_move'].tolist()

	for i in range(1,101):
		p_low = np.percentile(up_data, i)
		p_high = np.percentile(down_data, i)
		points_low = np.percentile(up_points, i)
		points_high = np.percentile(down_points, i)
		df_dict[names[j]]['percentile'].append(i)
		df_dict[names[j]]['Low_side'].append(round(p_low*100,2))
		df_dict[names[j]]['High_side'].append(round(p_high*100,2))
		df_dict[names[j]]['Points_UP'].append(round(points_low,2))
		df_dict[names[j]]['Points_Down'].append(round(points_high,2))


		# print(f'{i} Percentile : {p_low*100:.2f} % --- {p_high*100:.2f} %')

df_pct_daily = pd.DataFrame(df_dict['daily'])
df_pct_weekly = pd.DataFrame(df_dict['weekly'])
df_pct_monthly = pd.DataFrame(df_dict['monthly'])

# from_date = "2026-01-01"
# to_date   = "2026-01-06"

to_date_ = date.today()
from_date_ = to_date_ - timedelta(days=120)
from_date_long_ = to_date_ - timedelta(days=370)

to_date   = to_date_.strftime("%Y-%m-%d")
from_date = from_date_.strftime("%Y-%m-%d")
from_date_long = from_date_long_.strftime("%Y-%m-%d")


try:
	response_iv = api.get_historical_candle_data1(
					instrument_key="NSE_INDEX|India VIX",
					unit="days",
					interval="1",
					from_date=from_date_long,
					to_date=to_date)
	response_iv = response_iv.to_dict()
	df_iv = pd.DataFrame(response_iv['data']['candles'], columns=['time', 'open', 'high', 'low', 'close', 'vol1', 'vol2'])
	iv_data_close = df_iv['close']

	response = api.get_historical_candle_data1(
					instrument_key="NSE_INDEX|Nifty 50",
					unit="days",
					interval="1",
					from_date=from_date,
					to_date=to_date)
	response = response.to_dict()


	df = pd.DataFrame(response['data']['candles'], columns=['time', 'open', 'high', 'low', 'close', 'vol1', 'vol2'])
	df.drop(['vol1', 'vol2'], axis='columns', inplace=True)
	df['time'] = pd.to_datetime(df['time'])
	df.set_index('time', inplace=True)
	df = df.sort_index()
	previous_day_close = df['close'].iloc[-1]
	nif_close = df['close']

	log_return_20 = np.log(nif_close / nif_close.shift(1)).tail(20).dropna()
	std_20 = np.std(log_return_20, ddof=1)
	rv_20 = std_20*math.sqrt(252)*100

	log_return_30 = np.log(nif_close / nif_close.shift(1)).tail(30).dropna()
	std_30 = np.std(log_return_30, ddof=1)
	rv_30 = std_30*math.sqrt(252)*100

	# Weekly Candles
	weekly_df = df.resample('W-FRI').agg({'open': 'first', 'high': 'max', 'low':  'min', 'close': 'last'}).dropna()
	weekly_df.index = weekly_df.index - pd.to_timedelta(weekly_df.index.weekday, unit='D')
	last_week_df = weekly_df[:-1]
	previous_week_close = last_week_df['close'].iloc[-1]
	today = df.index.max().normalize()
	monday = today - pd.Timedelta(days=today.weekday())
	current_weekdays_df = pd.DataFrame(columns=['open', 'high', 'low', 'close'], dtype='float64')
	current_weekdays_df = df.loc[monday:today].copy()

	# Monthly candles
	monthly_df = df.resample('MS').agg({'open': 'first', 'high': 'max', 'low':  'min', 'close': 'last'}).dropna()
	monthly_df = monthly_df[:-1]
	previous_month_close = monthly_df['close'].iloc[-1]
	today = df.index.max().normalize()
	month_start = today.replace(day=1)
	current_monthdays_df = pd.DataFrame(columns=['open', 'high', 'low', 'close'], dtype='float64')
	current_monthdays_df = df.loc[month_start:today].copy()

	nifty_prev_close = response['data']['candles'][0][4]

except Exception as e:
	print("Exception:", e)

def plot_data():
	try:
		response = api.get_intra_day_candle_data("NSE_INDEX|Nifty 50", "minutes", "1")
		nifty_candle_data = response.to_dict()['data']['candles']
		nifty_candle_data = pd.DataFrame(nifty_candle_data, columns=('time', 'open', 'high', 'low', 'close', 'vol1', 'vol2')).drop(['vol1', 'vol2'], axis='columns')
		nifty_candle = nifty_candle_data['close'].iloc[::-1].reset_index(drop=True)


		response = api.get_intra_day_candle_data("NSE_INDEX|India VIX", "minutes", "1")
		vix_candle_data = response.to_dict()['data']['candles']
		vix_candle_data = pd.DataFrame(vix_candle_data, columns=('time', 'open', 'high', 'low', 'close', 'vol1', 'vol2')).drop(['vol1', 'vol2'], axis='columns')
		vix_candle = vix_candle_data['close'].iloc[::-1].reset_index(drop=True)

	except Exception as e:
		print("Exception:", e)

	return nifty_candle, vix_candle

# nifty_cand, india_vix_cand = plot_data()


current_date = today = pd.Timestamp.now(tz='Asia/Kolkata').normalize()

# Switch - line 197 to 223
# weekly_dte = 5
# monthly_dte = 25

def get_nifty_dte():
	inst_url = 'https://assets.upstox.com/market-quote/instruments/exchange/complete.csv.gz'
	inst = pd.read_csv(inst_url)

	expiries = inst[(inst['exchange'] == 'NSE_FO') & (inst['instrument_type'] == 'OPTIDX') &  (inst['name'] == 'NIFTY')]['expiry'].unique()
	expiries = pd.to_datetime(np.sort(expiries))
	today = pd.Timestamp.today().normalize()
	expiries = expiries[expiries >= today]

	# current month expiries
	curr_month_expiries = expiries[(expiries.month == expiries[0].month) & (expiries.year  == expiries[0].year)]
	next_expiry    = curr_month_expiries[0]
	monthly_expiry = curr_month_expiries[-1]

	# business days inclusive
	dte_next_expiry = np.busday_count(today.date(), (next_expiry + pd.Timedelta(days=1)).date())

	dte_monthly_expiry = np.busday_count(today.date(), (monthly_expiry + pd.Timedelta(days=1)).date())

	return [next_expiry, dte_next_expiry], [monthly_expiry, dte_monthly_expiry]

weekly, monthly = get_nifty_dte()

weekly_exp = weekly[0].date()
monthly_exp = monthly[0].date()
weekly_dte = weekly[1]
monthly_dte = monthly[1]

def pct_cal(prev_close, live_candle):
	low_points  = max(prev_close - live_candle['low'], 0)
	high_points = max(live_candle['high'] - prev_close, 0)
	curr_points = live_candle['close'] - prev_close

	low_pct  = low_points  / prev_close
	high_pct = high_points / prev_close
	curr_pct = curr_points / prev_close

	return low_pct, high_pct, curr_pct

plt.ion()
fig = plt.figure(figsize=(24, 13))
gs = gridspec.GridSpec(2, 3, height_ratios=[3, 2], hspace=0.25, wspace=0.22)


ax_d = fig.add_subplot(gs[0, 0])
ax_w = fig.add_subplot(gs[0, 1])
ax_m = fig.add_subplot(gs[0, 2])

# 4th plot (full width)
ax_price = fig.add_subplot(gs[1, :])

fig.subplots_adjust(left=0.05, right=0.96, top=0.95, bottom=0.07)


def init_full_dist_plot(ax, title, show_ylabel=False):
	# ---- Distribution line ----
	line, = ax.plot([], [], lw=1)

	# ---- Sigma lines ----
	s1p = ax.axvline(0, ls='--', lw=1, label='+1σ')
	s1n = ax.axvline(0, ls='--', lw=1, label='-1σ')

	s2p = ax.axvline(0, ls=':',  lw=1, alpha=0.7, label='+2σ')
	s2n = ax.axvline(0, ls=':',  lw=1, alpha=0.7, label='-2σ')

	s3p = ax.axvline(0, ls='-.', lw=1, alpha=0.5, label='+3σ')
	s3n = ax.axvline(0, ls='-.', lw=1, alpha=0.5, label='-3σ')

	# ---- Price lines ----
	low  = ax.axvline(0, color='red',  lw=1, label='Low')
	high = ax.axvline(0, color='red',  lw=1, label='High')
	curr = ax.axvline(0, color='blue', lw=1.5, label='Current')

	# ---- Axis formatting ----
	ax.set_title(title)
	ax.set_xlabel("Return")
	ax.grid(alpha=0.3)
	if show_ylabel:
		ax.set_ylabel("Probability Density")

	ax.legend(loc='upper right', fontsize=8)

	# ---- Sigma info box (top-left) ----
	sigma_box = ax.text(
		0.02, 0.95, "",
		transform=ax.transAxes,
		ha="left", va="top",
		fontsize=8,
		bbox=dict(facecolor="white", alpha=0.75, edgecolor="gray")
	)

	# ---- Text labels (σ + %) ----
	labels = {
		"low_sigma":  ax.text(0, 0, "", rotation=90, color="red", ha="right", va="center", fontsize=8),
		"low_pct":    ax.text(0, 0, "", rotation=90, color="red", ha="right", va="center", fontsize=8),

		"high_sigma": ax.text(0, 0, "", rotation=90, color="red", ha="left", va="center", fontsize=8),
		"high_pct":   ax.text(0, 0, "", rotation=90, color="red", ha="left", va="center", fontsize=8),

		"curr_sigma": ax.text(0, 0, "", rotation=90, color="blue", ha="left", va="center", fontsize=8),
		"curr_pct":   ax.text(0, 0, "", rotation=90, color="blue", ha="left", va="center", fontsize=8)}

	return {
		"ax": ax,
		"line": line,
		"s1p": s1p, "s1n": s1n,
		"s2p": s2p, "s2n": s2n,
		"s3p": s3p, "s3n": s3n,
		"low": low, "high": high, "curr": curr,
		"sigma_box": sigma_box,
		"labels": labels}

# Switch - line 318 to 320
# daily_plot = init_full_dist_plot(ax_d, f"Daily", show_ylabel=True)
# weekly_plot = init_full_dist_plot(ax_w, f"Weekly")
# monthly_plot = init_full_dist_plot(ax_m, f"Monthly")

daily_plot = init_full_dist_plot(ax_d, f"Daily: {date.today()}", show_ylabel=True)
weekly_plot = init_full_dist_plot(ax_w, f"Weekly: {weekly_exp} | DTE: {weekly_dte}")
monthly_plot = init_full_dist_plot(ax_m, f"Monthly: {monthly_exp}| DTE: {monthly_dte}")



def init_price_vix_plot(ax):
	ax.set_title("Nifty vs India VIX")
	ax.set_xlabel("Time")
	ax.set_ylabel("Nifty")
	ax.grid(alpha=0.3)

	# Second Y-axis for VIX
	ax_vix = ax.twinx()
	ax_vix.set_ylabel("India VIX")

	# Initialize lines
	nifty_line, = ax.plot([], [], color="black", lw=1.5, label="Nifty")
	vix_line,   = ax_vix.plot([], [], color="purple", lw=1.2, label="India VIX")

	# Combined legend
	lines_1, labels_1 = ax.get_legend_handles_labels()
	lines_2, labels_2 = ax_vix.get_legend_handles_labels()
	ax.legend(lines_1 + lines_2, labels_1 + labels_2,
			  loc="upper left", fontsize=8)

	return {
		"ax_price": ax,
		"ax_vix": ax_vix,
		"nifty_line": nifty_line,
		"vix_line": vix_line
	}

price_plot = init_price_vix_plot(ax_price)

ax_vix     = price_plot["ax_vix"]
nifty_line = price_plot["nifty_line"]
vix_line   = price_plot["vix_line"]


################################

plt.show(block=False)
plt.pause(0.1)

mng = plt.get_current_fig_manager()
mng.window.showMaximized()   # ✅ MAXIMIZED window (NOT fullscreen)

last_plot_time = 0

def update_plot(plot, sigma, low_pct, high_pct, curr_pct):
	mu = 0

	x = np.linspace(mu - 4*sigma, mu + 4*sigma, 1000)
	y = (1/(sigma*np.sqrt(2*np.pi))) * np.exp(-0.5*((x-mu)/sigma)**2)

	plot["line"].set_data(x, y)

	plot["s1p"].set_xdata([ sigma,  sigma])
	plot["s1n"].set_xdata([-sigma, -sigma])
	plot["s2p"].set_xdata([ 2*sigma,  2*sigma])
	plot["s2n"].set_xdata([-2*sigma, -2*sigma])
	plot["s3p"].set_xdata([ 3*sigma,  3*sigma])
	plot["s3n"].set_xdata([-3*sigma, -3*sigma])

	plot["low"].set_xdata([-low_pct,  -low_pct])
	plot["high"].set_xdata([ high_pct, high_pct])
	plot["curr"].set_xdata([ curr_pct, curr_pct])

	return x, y

def update_text(ax, x_range, labels,
				low_pct, high_pct, curr_pct,
				sigma_used):

	dx = (x_range.max() - x_range.min()) * 0.015  # slightly more horizontal space

	# FIXED vertical positions (AXES coords)
	y_pct   = 0.84   # TOP  → percentage
	y_sigma = 0.70   # BOTTOM → sigma (more gap now)

	# z-scores
	z_low  = low_pct  / sigma_used
	z_high = high_pct / sigma_used
	z_curr = curr_pct / sigma_used

	# ---------- LOW ----------
	labels["low_pct"].set_position((-low_pct - dx, y_pct))
	labels["low_pct"].set_text(f"{-low_pct*100:.2f}%")
	labels["low_pct"].set_transform(ax.get_xaxis_transform())

	labels["low_sigma"].set_position((-low_pct - dx, y_sigma))
	labels["low_sigma"].set_text(f"{-z_low:.2f}σ")
	labels["low_sigma"].set_transform(ax.get_xaxis_transform())

	# ---------- HIGH ----------
	labels["high_pct"].set_position((high_pct + dx, y_pct))
	labels["high_pct"].set_text(f"{high_pct*100:.2f}%")
	labels["high_pct"].set_transform(ax.get_xaxis_transform())

	labels["high_sigma"].set_position((high_pct + dx, y_sigma))
	labels["high_sigma"].set_text(f"{z_high:.2f}σ")
	labels["high_sigma"].set_transform(ax.get_xaxis_transform())

	# ---------- CURRENT ----------
	labels["curr_pct"].set_position((curr_pct + dx, y_pct))
	labels["curr_pct"].set_text(f"{curr_pct*100:.2f}%")
	labels["curr_pct"].set_transform(ax.get_xaxis_transform())

	labels["curr_sigma"].set_position((curr_pct + dx, y_sigma))
	labels["curr_sigma"].set_text(f"{z_curr:.2f}σ")
	labels["curr_sigma"].set_transform(ax.get_xaxis_transform())


last_price_plot_time = 0
PRICE_PLOT_INTERVAL = 90

app = xw.App(visible=True)
wb = app.books[0]   # default Book1

# delete default sheet

# add your sheets
for name in ['daily', 'weekly', 'monthly']:
	wb.sheets.add(name)

monthly_sheet = wb.sheets['monthly']
weekly_sheet  = wb.sheets['weekly']
daily_sheet   = wb.sheets['daily']
wb.sheets['Sheet1'].delete()

def update_excel(daily, weekly, monthly, df_ranges, india_vix):
	global daily_sheet, weekly_sheet, monthly_sheet

	daily_sheet.range("G1").options(index=False).value = df_ranges
	
	daily_sheet.range('A1').options(index=False).value = daily[3]
	daily_sheet.range('K1').value = 'Daily Live Candle'
	daily_sheet.range('K2').value = 'Low Level'
	daily_sheet.range('K3').value = 'High Level'
	daily_sheet.range('K4').value = 'Current Level'
	daily_sheet.range('K5').value = 'Time Period (Days)'
	daily_sheet.range('K6').value = 'India Vix'
	daily_sheet.range('L2').value = round(-daily[0]*100,2)
	daily_sheet.range('L3').value = round(daily[1]*100,2)
	daily_sheet.range('L4').value = round(daily[2]*100,2)
	daily_sheet.range('L6').value = india_vix

	weekly_sheet.range('A1').options(index=False).value = weekly[3]
	weekly_sheet.range('K2').value = 'Low Level'
	weekly_sheet.range('K3').value = 'High Level'
	weekly_sheet.range('K4').value = 'Current Level'
	weekly_sheet.range('L2').value = round(-weekly[0]*100,2)
	weekly_sheet.range('L3').value = round(weekly[1]*100,2)
	weekly_sheet.range('L4').value = round(weekly[2]*100,2)
	weekly_sheet.range('K1').value = 'Weekly Live Candle'

	monthly_sheet.range('A1').options(index=False).value = monthly[3]
	monthly_sheet.range('K2').value = 'Low Level'
	monthly_sheet.range('K3').value = 'High Level'
	monthly_sheet.range('K4').value = 'Current Level'
	monthly_sheet.range('L2').value = round(-monthly[0]*100,2)
	monthly_sheet.range('L3').value = round(monthly[1]*100,2)
	monthly_sheet.range('L4').value = round(monthly[2]*100,2)
	monthly_sheet.range('K1').value = 'Monthly Live Candle'

while True:
	rows = []
	with lock:
		today_candle = {'time':current_date, 'open':live_data['nifty']['open'], 'high':live_data['nifty']['high'], 'low':live_data['nifty']['low'], 'close':live_data['nifty']['close']}
		india_vix = live_data['india_vix']

	current_day_df = pd.DataFrame([today_candle])
	current_weekdays_df.loc[current_date] = [today_candle['open'], today_candle['high'], today_candle['low'], today_candle['close']]
	current_monthdays_df.loc[current_date] = [today_candle['open'], today_candle['high'], today_candle['low'], today_candle['close']]

	# Nifty - Daily Candle
	live_daily_candle = today_candle 
	daily_low_pct, daily_high_pct, daily_curr_pct = pct_cal(previous_day_close, live_daily_candle)
	daily_excel = [daily_low_pct, daily_high_pct, daily_curr_pct, df_pct_daily]

	india_vix_ltp = india_vix['close']
	daily_sigma = (india_vix_ltp/100)/math.sqrt(252)
	weekly_dte_sigma = daily_sigma * math.sqrt(weekly_dte)
	monthly_dte_sigma = daily_sigma * math.sqrt(monthly_dte)
	india_vix_change = (india_vix['close'] - india_vix['open']) / india_vix['open']
	india_vix_change_high = (india_vix['high'] - india_vix['open']) / india_vix['open']
	india_vix_change_low = (india_vix['low'] - india_vix['open']) / india_vix['open']

	# Nifty Weekly Candle
	live_weekly_candle = {'time':current_weekdays_df.index[0], 'open': current_weekdays_df['open'].iloc[0], 'high': current_weekdays_df['high'].max(), 'low': current_weekdays_df['low'].min(), 'close': current_weekdays_df['close'].iloc[-1]}
	weekly_low_pctt, weekly_high_pctt, weekly_curr_pctt = pct_cal(previous_week_close, live_weekly_candle)
	week_excel = [weekly_low_pctt, weekly_high_pctt, weekly_curr_pctt, df_pct_weekly]

	# Nifty Monthly Candle
	live_monthly_candle = {'time':current_monthdays_df.index[0], 'open': current_monthdays_df['open'].iloc[0], 'high': current_monthdays_df['high'].max(), 'low': current_monthdays_df['low'].min(), 'close': current_monthdays_df['close'].iloc[-1]}
	monthly_low_pctt, monthly_high_pctt, monthly_curr_pctt = pct_cal(previous_month_close, live_monthly_candle)
	month_excel = [monthly_low_pctt, monthly_high_pctt, monthly_curr_pctt, df_pct_monthly]

	weekly_low_pct   = daily_low_pct
	weekly_high_pct  = daily_high_pct
	weekly_curr_pct  = daily_curr_pct

	monthly_low_pct  = daily_low_pct
	monthly_high_pct = daily_high_pct
	monthly_curr_pct = daily_curr_pct

	now = time.time()

	if now - last_plot_time >= 10:

		if daily_sigma <= 0 or weekly_dte_sigma <= 0 or monthly_dte_sigma <= 0:
			last_plot_time = now
			plt.pause(0.05)
			continue

		
		time_period = daily_sheet.range("L5").value
		if time_period == None:
			time_period = 1

		for c in range(1, 100):   # 1% to 99% (100% is infinite)
			z = norm.ppf((1 + c/100) / 2)   # two-sided confidence
			rng = z * daily_sigma * math.sqrt(time_period)
			rows.append({
				"confidence_%": c,
				"z_score": round(z, 2),
				"range_%": round(rng*100, 2)})

		df_ranges = pd.DataFrame(rows)

		update_excel(daily_excel, week_excel, month_excel, df_ranges, india_vix_ltp)

		ivp = (iv_data_close <= india_vix_ltp).sum() / len(iv_data_close) * 100

		if last_price_plot_time == 0 or now - last_price_plot_time >= PRICE_PLOT_INTERVAL:

			nifty_cand, india_vix_cand = plot_data()

			nifty_line.set_data(nifty_cand.index, nifty_cand)
			vix_line.set_data(india_vix_cand.index, india_vix_cand)

			ax_price.relim()
			ax_price.autoscale_view()

			ax_vix.relim()
			ax_vix.autoscale_view()

			last_price_plot_time = now
	
	
		nifty_cand, india_vix_cand = plot_data()

		nifty_line.set_data(nifty_cand.index, nifty_cand)

		vix_line.set_data(india_vix_cand.index, india_vix_cand)

		# Autoscale both axes independently
		ax_price.relim()
		ax_price.autoscale_view()

		ax_vix.relim()
		ax_vix.autoscale_view()

		vix_ltp = india_vix_ltp * 1.0   # just clarity
		vix_dh  = india_vix_change_high * 100
		vix_dc  = india_vix_change * 100
		vix_dl  = india_vix_change_low * 100

		daily_plot["sigma_box"].set_text(f"Sigma\n{daily_sigma*100:.2f}%\n\nVIX: {vix_ltp:.2f}\nRV20 :{rv_20:.2f}\nRV30: {rv_30:.2f}\n\nΔH {vix_dh:+.2f}%\nΔC {vix_dc:+.2f}%\nΔL {vix_dl:+.2f}%\n\nIVP: {ivp:.2f}")
		weekly_plot["sigma_box"].set_text(f"Sigma\n{weekly_dte_sigma*100:.2f}%\n\nVIX: {vix_ltp:.2f}\nRV20 :{rv_20:.2f}\nRV30: {rv_30:.2f}\n\nΔH {vix_dh:+.2f}%\nΔC {vix_dc:+.2f}%\nΔL {vix_dl:+.2f}%\n\nIVP: {ivp:.2f}")
		monthly_plot["sigma_box"].set_text(f"Sigma\n{monthly_dte_sigma*100:.2f}%\n\nVIX: {vix_ltp:.2f}\nRV20 :{rv_20:.2f}\nRV30: {rv_30:.2f}\n\nΔH {vix_dh:+.2f}%\nΔC {vix_dc:+.2f}%\nΔL {vix_dl:+.2f}%\n\nIVP: {ivp:.2f}")

		x_d, y_d = update_plot(daily_plot, daily_sigma,	daily_low_pct, daily_high_pct, daily_curr_pct)
		x_w, y_w = update_plot(weekly_plot, weekly_dte_sigma, weekly_low_pct,weekly_high_pct, weekly_curr_pct)
		x_m, y_m = update_plot(monthly_plot, monthly_dte_sigma,	monthly_low_pct, monthly_high_pct, monthly_curr_pct)

		update_text(daily_plot["ax"], x_d, daily_plot["labels"], daily_low_pct, daily_high_pct, daily_curr_pct, daily_sigma)
		update_text(weekly_plot["ax"], x_w, weekly_plot["labels"], weekly_low_pct, weekly_high_pct, weekly_curr_pct, weekly_dte_sigma)
		update_text(monthly_plot["ax"], x_m, monthly_plot["labels"], monthly_low_pct, monthly_high_pct, monthly_curr_pct, monthly_dte_sigma)

		ax_d.set_xlim(x_d.min(), x_d.max())
		ax_d.set_ylim(0, y_d.max() * 1.25)

		ax_w.set_xlim(x_w.min(), x_w.max())
		ax_w.set_ylim(0, y_w.max() * 1.25)

		ax_m.set_xlim(x_m.min(), x_m.max())
		ax_m.set_ylim(0, y_m.max() * 1.25)

		fig.canvas.draw_idle()

		last_plot_time = now

	exit = daily_sheet.range('L7').value
	if exit == 'e':
		streamer.disconnect()
		plt.close('all')
		break

	# time.sleep(1)
	plt.pause(0.05)
	