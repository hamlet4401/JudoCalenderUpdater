import Spond.spond.base
from Spond.spond import spond
import aiohttp
import os
import json
import getpass

GROUP_ID = 'E121119905434844A916B218A1B82A3A'  # Medewerkers groep
SUBGROUP_ID = 'DC3CB534C6F7436DA49258A345DA1CE2'  # Lestabel subgroep

SPOND_CREDENTIALS = "credentials/spond_credentials.json"


async def fetch_spond_credentials():
    """Fetch Spond credentials from the specified JSON file."""
    try:
        with open(SPOND_CREDENTIALS, 'r') as f:
            credentials = json.load(f)
            username = credentials["username"]
            password = credentials["password"]
            while not await perform_credentials_check(username, password):
                username, password = ask_spond_credentials(repeated=True)
    except (FileNotFoundError, KeyError):
        username, password = ask_spond_credentials()
        while not await perform_credentials_check(username, password):
            username, password = ask_spond_credentials(repeated=True)
    return username, password


async def perform_credentials_check(username, password):
    """Check if the provided Spond credentials are valid."""
    try:
        s = spond.Spond(username=username, password=password)
        groups = await s.get_groups()
        return True
    except Spond.spond.base.AuthenticationError as e:
        print()
        print(e)
        print()
        return False


def ask_spond_credentials(repeated=False):
    """Prompt the user for Spond credentials."""
    if repeated:
        error_message = "Username and or password are incorrect. Please try again."
        print()
        print(error_message)
        print()
    username = input("Enter your Spond username > ")
    password = getpass.getpass("Enter your Spond password > ")
    print()
    with open(SPOND_CREDENTIALS, 'w') as f:
        json.dump({"username": username, "password": password}, f)
    return username, password


async def download_media(url: str, filename: str, tmp_dir_name: str):
    file_path = os.path.join(tmp_dir_name, filename)

    async with aiohttp.ClientSession() as session:
        async with session.get(url) as response:
            if response.status == 200:
                with open(file_path, 'wb') as f:
                    f.write(await response.read())
                print(f"File downloaded and saved to {file_path}")
            else:
                print(f"Failed to download file. Status code: {response.status}")

    return file_path


async def get_latest_time_table(tmp_dir_name):
    username, password = await fetch_spond_credentials()
    s = spond.Spond(username=username, password=password)
    posts = await s.get_posts_for_group(group_id=GROUP_ID, subgroup_id=SUBGROUP_ID, max_posts=1)

    file_path = ""

    # Assuming we want to download the first attachment
    if posts and posts[0]["attachments"]:
        media = posts[0]["attachments"][0]["media"]
        filename = posts[0]["attachments"][0]["title"] + ".xlsx"

        # Download the media to a temporary folder
        file_path = await download_media(media, filename, tmp_dir_name)
        print(f"Downloaded file is located at: {file_path}")

    await s.clientsession.close()

    return file_path
