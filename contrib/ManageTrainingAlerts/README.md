# Create and delete training alerts

A python script to easily create a set of 'dummy' alerts, and then delete them afterwards.
Useful during a training session of TheHive.

# Usage

## Create new random alerts

```
training-alert.py create 5
```

## Delete one alert

```
training-alert.py delete _id 1fd272233688b6cd685b138092970ce8
```

Uses the **force** parameter to complete delete the alert.

## Delete all alerts with a specific tag

```
 training-alert.py delete tag Perimeter
 ```

Calls the "delete one alert" recursively. Range is set to all.

