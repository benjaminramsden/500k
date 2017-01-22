import sys
from imgurpython import ImgurClient
import os
import zipfile

def main(argv=None):

    # Authenticate the user here, authentication details go in auth.ini.
    config = get_config()
    config.read('auth.ini')
    client_id = config.get('credentials', 'client_id')
    client_secret = config.get('credentials', 'client_secret')
    access_token = config.get('credentials', 'access_token')
    refresh_token = config.get('credentials', 'refresh_token')

    if access_token is None or refresh_token is None:
        client = authenticate()
    else:
        client = ImgurClient(client_id,
                             client_secret,
                             access_token,
                             refresh_token)

    # Get the dropbox directory
    dropbox_dir = os.getcwd()

    # Rename files to *.zip
    [os.rename(f, f.replace('.docx', '.zip')) for f in os.listdir(dropbox_dir) if not f.startswith('.')]

    # unzip docx (now zip) files
    for file in os.listdir(dropbox_dir):
        if file.endswith('.zip'):
            zip_ref = zipfile.ZipFile(dropbox_dir + "/" + file, 'r')
            zip_ref.extractall(dropbox_dir)
            zip_ref.close()

    # For posting:
    # - Create monthly album (if doesn't already exist) use dropbox folder name
    # - Post all profile pics to that album
    dropbox_dir = dropbox_dir.rpartition('/')[2]
    print "dropbox_dir: {0}".format(dropbox_dir)

    # Check Imgur to see if this album already exists
    album_ids = client.get_account_album_ids('ramalamman', page=0)

    found_dir = False
    for album in album_ids:
        album_dir = client.get_album(album)
        if album_dir.title == dropbox_dir:
            album_id = album
            found_dir = True
            break
    if not found_dir:
        album_id = client.create_album(fields={'title': dropbox_dir, 'privacy': 'Hidden'})

    # Post all pictures from unzipped directory into this album
    # Requires uploading image first with album to assign to
    for dirName, subdirList, fileList in os.walk(dropbox_dir):
        print('Found directory: %s' % dirName)
        for fname in fileList:
            print('\t%s' % fname)
            if fname.lower().endswith(('.png', '.jpg', '.jpeg')):
                config = {
                    'album': album_id,
                    'name': fname,
                    'title': fname
                }
                client.upload_from_path(path=dirName + "/" + fname,
                                        config=config,
                                        anon=False)
                print "File {0} uploaded".format(fname)

    return 0        # success

def authenticate():
    # Get client ID and secret from auth.ini
    config = get_config()
    config.read('auth.ini')
    client_id = config.get('credentials', 'client_id')
    client_secret = config.get('credentials', 'client_secret')

    client = ImgurClient(client_id, client_secret)

    # Authorization flow, pin example (see docs for other auth types)
    authorization_url = client.get_auth_url('pin')
    #print authorization_url

    #pin = authorization_url.rpartition('=')[2]
    #print "PIN: {0}".format(pin)

    print("Go to the following URL: {0}".format(authorization_url))

    # Read in the pin, handle Python 2 or 3 here.
    pin = get_input("Enter pin code: ")

    # redirect user to `authorization_url`, obtain pin (or code or token)
    credentials = client.authorize(pin, 'pin')
    client.set_user_auth(credentials['access_token'],
                         credentials['refresh_token'])

    print("Authentication successful! Here are the details:")
    print("   Access token:  {0}".format(credentials['access_token']))
    print("   Refresh token: {0}".format(credentials['refresh_token']))

    config.set('credentials', 'refresh_token', credentials['refresh_token'])
    config.set('credentials', 'access_token', credentials['access_token'])
    with open('auth.ini', 'wb') as configfile:
        config.write(configfile)

    return client


def get_input(string):
    ''' Get input from console regardless of python 2 or 3 '''
    try:
        return raw_input(string)
    except:
        return input(string)

def get_config():
    ''' Create a config parser for reading INI files '''
    try:
        import ConfigParser
        return ConfigParser.ConfigParser(allow_no_value=True)
    except:
        import configparser
        return configparser.ConfigParser(allow_no_value=True)

def rename(dir, pattern, titlePattern):
    for pathAndFilename in glob.iglob(os.path.join(dir, pattern)):
        title, ext = os.path.splitext(os.path.basename(pathAndFilename))
        os.rename(pathAndFilename,
                  os.path.join(dir, titlePattern % title + ext))

if __name__ == '__main__':
    status = main()
    sys.exit(status)
