#!/usr/bin/env python3
# -*- coding: utf-8 -*-

# Author : Francois Ihry (github:francoisihry)

import argparse
import requests
import json


def insert_datatypes(thehive_url, api_key, datatypes):
    for dt in datatypes:
        response = requests.post('{}/api/list/list_artifactDataType'.format(thehive_url),
                                 headers={"Authorization": "Bearer {}".format(api_key)},
                                 data={"value": dt})
        if response.status_code == 400:
            print(response.text)
        elif response.status_code != 200:
            raise Exception("{} : {}".format(response.status_code, response.text))


def insert_case_templates(thehive_url, api_key, case_templates):
    for ct in case_templates:
        data = {k:ct[k] for k in ct.keys() if k in["titlePrefix", "customFields", "metrics",
                                                   "description","name", "titlePrefix", "severity",
                                                   "tlp", "status", "tasks", "tags"]}
        response = requests.post('{}/api/case/template'.format(thehive_url),
                                 headers={"Authorization": "Bearer {}".format(api_key),
                                          'Content-Type': 'application/json'},
                                 data=json.dumps(data))
        if response.status_code == 400:
            print(response.text)
        elif response.status_code != 201:
            raise Exception("{} : {}".format(response.status_code, response.text))


def insert_custom_fields(thehive_url, api_key, custom_fields):
    for cf in custom_fields:
        response = requests.post('{}/api/list/custom_fields'.format(thehive_url),
                                 headers={"Authorization": "Bearer {}".format(api_key),
                                          'Content-Type': 'application/json'},
                                 data=json.dumps({"value": cf}))
        if response.status_code == 400:
            print(response.text)
        elif response.status_code != 200:
            raise Exception("{} : {}".format(response.status_code, response.text))


def insert_metrics(thehive_url, api_key, metrics):
    for m in metrics:
        response = requests.post('{}/api/list/case_metrics'.format(thehive_url),
                                 headers={"Authorization": "Bearer {}".format(api_key),
                                          'Content-Type': 'application/json'},
                                 data=json.dumps({"value": m}))
        if response.status_code == 400:
            print(response.text)
        elif response.status_code != 200:
            raise Exception("{} : {}".format(response.status_code, response.text))


def submit_config(thehive_url, api_key, conf_path):
    with open(conf_path, 'r') as infile:
        config = json.loads(infile.read())
        print("Inserting observable  datatypes...")
        insert_datatypes(thehive_url, api_key, config["datatypes"])
        print("Inserting custom fields...")
        insert_custom_fields(thehive_url, api_key, config["custom_fields"])
        print("Inserting metrics...")
        insert_metrics(thehive_url, api_key, config["metrics"])
        print("Inserting case templates...")
        insert_case_templates(thehive_url, api_key, config["case_templates"])

def parse_args():
    parser = argparse.ArgumentParser()
    parser.add_argument("-k","--key",
        help=("API key"),required=True)
    parser.add_argument("-u","--url",
        help=("TheHive server url."),required=True)
    parser.add_argument("-c", "--config_path",
                        help=("Configuration file path."), required=True)
    return parser.parse_args()


def main():
    args = parse_args()
    print("Submitting config to  {}".format(args.url))
    submit_config(args.url, args.key, args.config_path)


if __name__ == "__main__" :
    main()
