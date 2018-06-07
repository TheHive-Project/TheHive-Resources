# Export/Import TheHive config

These script allow to
    - retrieve the configuration (datatypes, case templates, metrics and custom fields) and store it in a file.
    - submit a configuration to an instance of theHive



# How to use:

- Get the configuration:

```
./fetch_config.py -k <API_KEY> -u <theHive url> -o /tmp/my_config.conf
```

- Submit the configuration:
```
./fetch_config.py -k <API_KEY> -u <theHive url> -o /tmp/my_config.conf
```
 ConflictError messages can be printed if one element already exists (but wont stop the process)