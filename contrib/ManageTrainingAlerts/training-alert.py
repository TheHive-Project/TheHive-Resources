#!/usr/bin/env python3

'''
TheHive - Create and Delete alerts. Useful during a training session.
Koen Van Impe - 2020

 Create 5 alerts:
    training-alert.py create 5
 Delete specific alert:
    training-alert.py delete _id 1fd272233688b6cd685b138092970ce8
 Delete alerts with a tag:
    training-alert.py delete tag Perimeter
'''

import requests
import json
import random
import sys

auth = "<AUTHKEY>"
host = "http://127.0.0.1:9000"


def create_alert():
    maxrand = 10000
    sources = ['IDS', 'AV', 'Firewall', 'Honeypot']
    alert_type = ['Internal', 'External', 'Human']
    title = ['Rare process', 'Rare scheduled task', 'Unusual activity', 'Outbound tunnel', 'Cleartext traffic', 'Malware alert', 'New SUID']
    title_ip = '{}.{}.{}.{}'.format(random.randrange(1, 223), random.randrange(1, 223), random.randrange(1, 223), random.randrange(1, 223))
    tags = ['MISP', 'Sigma', 'Perimeter', 'BIA:1', 'High-confidence']

    headers = {'Content-Type': 'application/json', 'Authorization': 'Bearer {}'.format(auth), 'Accept': 'text/plain'}
    url = "{}/api/alert".format(host)

    data = {'title': '{} - {}'.format(random.choice(title), title_ip), 'description': 'Alert Description', 'type': random.choice(alert_type), 'source': random.choice(sources), 'sourceRef': '{} - {} - {}'.format(random.randrange(maxrand), random.randrange(maxrand), random.randrange(maxrand)), 'tags': [random.choice(tags), random.choice(tags)], 'artifacts': [{'dataType': 'ip', 'data': title_ip, 'message': 'Victim'}]}

    result = requests.post(url, headers=headers, data=json.dumps(data))
    if result.json()['status'] == 'New':
        print('Alert {} added'.format(result.json()['_id']))
    else:
        print('Failed to add alert')
        print(data)


def delete_alert(_id=False, tag=False):
    if _id:
        url = "{}/api/alert/{}?force=1".format(host, _id)
        headers = {'Authorization': 'Bearer {}'.format(auth)}
        result = requests.delete(url, headers=headers)
        if result.status_code == 204:
            print('Alert {} deleted'.format(_id))
        else:
            print('Failed to delete {}'.format(_id))

    if tag:
        url = "{}/api/alert/_search?range=all".format(host)
        headers = {'Content-Type': 'application/json', 'Authorization': 'Bearer {}'.format(auth), 'Accept': 'text/plain'}
        data = {'query': {'tags': tag}}
        result = requests.post(url, headers=headers, data=json.dumps(data))
        if result.status_code == 200:
            if result.json():
                for alert in result.json():
                    delete_alert(_id=alert['_id'])
            else:
                print("Nothing to delete")
        else:
            print("Failed to delete", result)


if len(sys.argv) > 1:
    action = sys.argv[1]
    if action == "create":
        i = 0
        count = int(sys.argv[2])
        while i <= count:
            create_alert()
            i = i + 1
    if action == "delete":
        subaction = sys.argv[2]
        if subaction == "_id":
            _id = sys.argv[3]
            delete_alert(_id=_id)
        elif subaction == "tag":
            tag = sys.argv[3]
            delete_alert(tag=tag)
else:
    print("Invalid arguments")
