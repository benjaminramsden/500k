import sys
from imgurpython import ImgurClient
import os
import zipfile
from sheets_api import update_sheet, get_all_missionary_reports

def main(argv=None):

    # Authenticate the user here, authentication details go in auth.ini.
    client = start_client()

    return post_images(client)

def start_client():
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
    return client

def post_images(client):
    # Get the dropbox parent directory containing the factfiles
    dropbox_dir = "C:\Users\\br1\Dropbox\NCM\\500k advocates & Missionary factfiles\New missionary factfiles by month"

    # Walk the subdirectories, copying each doc and docx file and replacing
    # the file extension with .zip and extract
    for dirpath, dirnames, filenames in os.walk("."):
        for filename in [f for f in filenames if f.endswith((".doc",".docx"))]:
            print("Extracting image from file: {}".format(filename))
            new_f = filename.split(".")[0]+".zip"
            shutil.copyfile(filename,new_f)
            zip_ref = zipfile.ZipFile(new_f, 'r')
            zip_ref.extractall(dirpath)
            zip_ref.close()

    # Post these pictures using the missionary ID as the title.
    # Due to limitations of the previous format, change the filename manually
    # on Imgur.
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

def update_imgur_ids():
    # Have the image ID from Imgur put into the spreadsheet against each
    # report so we know where to grab it from.
    client = start_client()

    # Get image data, match title of file to Missionary ID and then add this
    # data to the end of the list.
    (miss_ids, imgur_ids) = get_all_missionary_reports(imgur=True)
    images = client.get_account_images('me')

    # By default Google does not keep empty list entries at the end of a
    # row/column, correct this now
    imgur_ids.extend([[]]*(len(miss_ids)-len(imgur_ids)))

    # WARNING - This loop is slow, should try to speed it up
    for idx,imgur_id in enumerate(imgur_ids):
        if not imgur_id:
            # This needs updating, find the Imgur picture based on it's title
            # which should match the Missionary ID of the same index
            for image in images:
                if image.title==miss_ids[idx]:
                    # Matched, update the Imgur ID list
                    print("Updating photo for Missionary ID {0}, the image ID"
                          "is {1}".format(image.title,image.id))
                    imgur_id = image.id

    # TODO - Upload the new data to the spreadsheet
    update_sheet(imgur_ids,imgur=True)

if __name__ == '__main__':
    status = main()
    sys.exit(status)
