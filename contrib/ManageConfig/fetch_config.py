#!/usr/bin/env python3
# -*- coding: utf-8 -*-

# Author : Francois Ihry (github:francoisihry)

import argparse
import requests
import json
import os

def get_data_types(thehive_url, api_key):
    response = requests.get("{}/api/list/list_artifactDataType".format(thehive_url),
                            headers={"Authorization": "Bearer {}".format(api_key),
                                     'Content-Type': 'application/json'})
    if response.status_code != 200:
        raise Exception("{} : {}".format(response.status_code, response.text))
    return list(response.json().values())


def get_case_templates(thehive_url, api_key):
    response = requests.post("{}/api/case/template/_search".format(thehive_url),
                            headers={"Authorization": "Bearer {}".format(api_key)},
                             params={'range':'all'})
    if response.status_code != 200:
        raise Exception("{} : {}".format(response.status_code, response.text))
    return response.json()


def get_custom_fields(thehive_url, api_key):
    response = requests.get("{}/api/list/custom_fields".format(thehive_url),
                            headers={"Authorization": "Bearer {}".format(api_key),
                                     'Content-Type': 'application/json'})
    if response.status_code != 200:
        raise Exception("{} : {}".format(response.status_code, response.text))
    return list(response.json().values())


def get_metrics(thehive_url, api_key):
    response = requests.get("{}/api/list/case_metrics".format(thehive_url),
                            headers={"Authorization": "Bearer {}".format(api_key),
                                     'Content-Type': 'application/json'})
    if response.status_code != 200:
        raise Exception("{} : {}".format(response.status_code, response.text))
    return list(response.json().values())


def parse_args():
    parser = argparse.ArgumentParser()
    parser.add_argument("-k","--key",
        help=("API key"),required=True)
    parser.add_argument("-u","--url",
        help=("TheHive server url."),required=True)
    parser.add_argument("-o", "--output",
                        help=("Output path to store the config"))
    return parser.parse_args()


def fetch_config(thehive_url, api_key, output):
    if output==None:
        output="fetched_config.conf"
    elif os.path.isdir(output):
        output = os.path.join(output,"fetched_config.conf")
    output = os.path.join(os.path.dirname(os.path.realpath(__file__)), output)
    datatypes = get_data_types(thehive_url, api_key)
    case_templates = get_case_templates(thehive_url, api_key)
    custom_fields = get_custom_fields(thehive_url, api_key)
    metrics = get_metrics(thehive_url, api_key)
    config = {"datatypes": datatypes,
              "case_templates": case_templates,
              "custom_fields":custom_fields,
              "metrics":metrics}
    print("Saving config to {}".format(output))
    with open(output, 'w') as outfile:
        json.dump(config, outfile)


def main():
    args = parse_args()
    print("Fetching config from {}".format(args.url))
    fetch_config(args.url, args.key, args.output)


if __name__ == "__main__" :
    main()
