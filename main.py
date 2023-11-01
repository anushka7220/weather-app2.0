import pandas as pd
import json
import requests
import win32com.client as wincom
from sklearn.ensemble import RandomForestClassifier

# Create a speech object
speak = wincom.Dispatch("SAPI.SpVoice")


def get_weather_data(city):
    url = f"https://api.weatherapi.com/v1/current.json?key=d0501c165e5348cc938155442231408&q={city}"
    response = requests.get(url)
    if response.status_code == 200:
        weather_data = response.json()
        temperature = weather_data["current"]["temp_c"]
        humidity = weather_data["current"]["humidity"]
        precip_mm = weather_data["current"]["precip_mm"]
        is_raining = 1.0 if precip_mm > 0.05 else 0.0
        # Create a dictionary to store the data
        data = {'Temperature': [temperature], 'Humidity': [humidity], 'IsRaining': [is_raining]}
        print(data)
        # Create a DataFrame from the dictionary
        df = pd.DataFrame(data)
        return temperature, humidity, is_raining, df
    else:
        print(f"Failed to retrieve weather data for {city}. Status code: {response.status_code}")
        return None, None, None, None


# Create a Random Forest model (you can adjust hyperparameters as needed)
model = RandomForestClassifier(n_estimators=100, random_state=42)

# Make predictions on new weather data
city: str = input('Enter the name of the city: \n')
temperature, humidity, is_raining, weather_df = get_weather_data(city)
if temperature is not None and humidity is not None:
    X_train = weather_df[['Temperature', 'Humidity']]
    y_train = weather_df['IsRaining']
    # Train the logistic regression model
    model.fit(X_train, y_train)
    # Make predictions using the model
    new_data = [[temperature, humidity]]
    raining_prediction = model.predict(new_data)

    if raining_prediction[0] == 1:
        prediction = f"It's likely raining in {city}"
    else:
        prediction = f"It's probably not raining in {city}"

    text = f"The current weather in {city} is {temperature} degrees and {prediction}"
    print(text)
    speak.Speak(text)
