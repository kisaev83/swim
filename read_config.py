from configparser import ConfigParser

# /usr/local/bin/beswimbot/config.ini
# C:/Users/User/PycharmProjects/swimming/config.ini
def read_db_config(filename='C:/Users/User/PycharmProjects/swimming/config.ini', section=None):

    # create parser and read ini configuration file

    parser = ConfigParser()
    parser.read(filename)
    db = {}
    if parser.has_section(section):
        items = parser.items(section)
        for item in items:
            db[item[0]] = item[1]
    else:
        raise Exception('{0} not found in the {1} file'.format(section, filename))

    return db