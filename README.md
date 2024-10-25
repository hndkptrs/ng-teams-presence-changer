# Non-Graph Teams Presence Changer

Ever wanted to change the Teams presence on your company issued Teams through code, just to find out you need access to the Microsoft Graph api to do so? 

Worry not, by using the Teams web api, we can go ahead and adjust these settings without access to the Graph api. \
This is especially useful when using e.g. the webclient of Teams on Linux, which can't recognise you're online unless you are triggering a mouse-over event. 

This method does come with some drawbacks though, first and foremost we need to provide a fresh bearer token everyday and for some features, like set_presence, you need to input a valid session uuid at least once. 

## How to get your bearer token and session uuid
1. Open the developer menue on your browser.

1. Switch to the network tab, enable capturing and load the Teams website.

1. Search for "endpoints".

1. Get the bearer token from the request headers of this request.

1. Now switch to the Payload section of the request. The id mentioned within the json is your session uuid.
## Examples

```python
from TeamsPresenceChanger import *

BEARER = "INSERT BEARER HERE, ommit the BEARER keyword"

teams = TeamsPresenceChanger(BEARER,"INSERT SESSION ID HERE")

print(teams.get_status())
teams.set_presence("Busy","InACall", 60)
teams.set_publishnote("Try again later!", False)
teams.set_worklocation("remote", "2024-12-30T23:00:00.000Z")
teams.set_forced_availability("Offline", "2024-12-30T23:00:00.000Z")
```

### Your company has blocked Linux through conditional access?
No problem, just add the user agent of an os, that is allowed while creating the instance of your object.
```python
BEARER = "INSERT BEARER HERE, ommit the BEARER keyword"
USER_AGENT = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/130.0.0.0 Safari/537.36 Edg/129.0.2792.79"
teams = TeamsPresenceChanger(BEARER,"INSERT SESSION ID HERE", user_agent = USER_AGENT)
```

### Disclaimer
This doesn't work for the private version of Microsoft Teams.