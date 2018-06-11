# Export/Import TheHive config

These script allow to
    - retrieve the configuration (datatypes, case templates, metrics and custom fields) and store it in a file.
    - submit a configuration to an instance of theHive

# Config file format:

The configuration file should contain a json structure like below : 

```
{
    "datatypes": <observable data types>,
    "case_templates": <case templates>,
    "custom_fields": <custom fields>,
    "metrics": <metrics>
}
```
Take a look at [Create a configuration file](#create-a-configuration-file) which is more explicit.

# How to use:

## Get the configuration:

```
./fetch_config.py -k <API_KEY> -u <theHive url> -o /tmp/my_config.conf
```

## Create a configuration file: 

```
echo '
{
    "datatypes": ["url", "other", "user-agent", "regexp", "mail_subject", "registry", "mail", "autonomous-system", "domain", "ip", "uri_path", "filename", "hash", "file", "fqdn"],
    "case_templates": [
        {"name": "Vulnerability", "description": "When an import vulnerability is discovered a case should be created to analyze the compromise of the information system.",
        "customFields": {"cVE": {"string": null, "order": 1}},
        "tasks": [{"order": 0, "title": "Analyze the compromise", "description": "take a look at the information system architecture and softwares"}],
        "titlePrefix": "vul",
        "metrics": {"score": null},
        "severity": 2,
        "tlp": 2,
        "tags": []}
    ],
    "custom_fields": [
        {"name": "CVE",
        "reference": "cVE",
        "description": "CVE id of the vulnerability",
        "type": "string",
        "options": []}
    ],
    "metrics": [
        {"name": "score",
        "title": "Score",
        "description": "score of the vulnerability"}
    ]
}
'>/tmp/my_config.conf
```

## Submit the configuration:
```
./submit_config.py -k <API_KEY> -u <theHive url> -c /tmp/my_config.conf
```
 ConflictError messages can be printed if one element already exists (but wont stop the process)
 
 
# Dependences :

Python3
