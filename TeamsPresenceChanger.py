import requests
import time

class TeamsPresenceChanger:
    base_url = "https://presence.teams.microsoft.com/v1/me"
    user_agent = 'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/129.0.0.0 Safari/537.36'
    session_uuid = ""

    def __init__(self, bearer:str, session_uuid=session_uuid, base_url=base_url, user_agent=user_agent, timeout = 10):
        self.base_url = base_url
        self.generate_header(bearer, user_agent)
        self.set_session_uuid(session_uuid)
        self.timeout = timeout

    def generate_header(self, bearer:str, user_agent:str):
        self.headers = {
            "User-Agent": user_agent,
            "Authorization": f"Bearer {bearer}",
            "Content-Type" : "application/json" 
        }

    def set_session_uuid(self, session_uuid:str):
        self.session_uuid = session_uuid

    def set_forced_availability(self, availability:str, expiry="9999-12-30T23:00:00.000Z") -> str:
        """
        This functions sets the forced availability.

        Input: 
        [String] availability
          Possible values:
            - Available
            - Busy 
            - DoNotDistrub 
            - BeRightBack 
            - Away 
            - Offline

        Output: 
        [String] response
        """
        if (availability == "Available" 
            or availability == "Busy" 
            or availability == "DoNotDistrub" 
            or availability == "BeRightBack" 
            or availability == "Away" 
            or availability == "Offline"):
            url = self.base_url+"/forceavailability"
            payload = {
                "availability": availability
            }
            if availability == "Offline":
                payload.update({"activity": "OffWork"})
            if expiry:
                payload.update({"desiredExpirationTime": expiry})
            response = requests.put(url=url, json=payload, headers=self.headers, timeout=self.timeout)
            if response.status_code >=20 and response.status_code <=200:
                return response
            else:
                raise Exception(response)
        else:
            raise ValueError(f"Value availability {availability} is incorrect, possible values: \
            Available, Busy, DoNotDistrub, BeRightBack, BeRightBack, Away, Offline")

    def set_publishnote(self, note:str, pinned=True, expiry="9999-12-30T23:00:00.000Z") -> str:
        """
        This function sets the publishnote for your account.

        Input:
        [String] note
        [Bool] pinned
        [String] expiry
          Format: yyyy-MM-ddTHH:mm:ss.SSSZ
        
        Output: 
        [String] response
        """
        url = self.base_url+"/publishnote"
        message = f"<p>{note}</p>"
        if pinned:
            message += "<pinnednote></pinnednote>"
        payload = {
            "expiry": expiry,
            "message": message
        }        
        response = requests.put(url=url, json=payload, headers=self.headers, timeout=self.timeout)
        if response.status_code >=20 and response.status_code <=200:
            return response
        else:
            raise Exception(response)

    def set_worklocation(self, location:str, expiry="9999-12-30T23:00:00.000Z") -> str:
        """
        This function sets your worklocation to remote, office.

        Input:
        [String] location
          Possible values:
            - remote
            - office
            - reset
        [String] expiry
          Format: yyyy-MM-ddTHH:mm:ss.SSSZ

        Output: 
        [String] response
        """
        if location == "remote":
            location_int = 2
        elif location == "office":
            location_int = 1
        elif location == "reset":
            location_int = 0
        else:
            raise ValueError("Location must have the value remote, office or reset.")
        url = self.base_url+"/workLocation"
        payload = {
            "location": location_int,
            "expirationTime": expiry
        }
        response = requests.put(url=url, json=payload, headers=self.headers, timeout=self.timeout)
        if response.status_code >=20 and response.status_code <=200:
            return response
        else:
            raise Exception(response)

    def get_status(self) -> dict:
        """This function returns a dict with the current state of your teams profile."""
        url = self.base_url+"/presence"
        response = requests.get(url=url, headers=self.headers, timeout=self.timeout)
        if response.status_code >=20 and response.status_code <=200:
            return response.json()
        else:
            raise Exception(response)

    def set_presence(self, availability:str, activity:str, duration_in_m=240) -> str:
        """
        This function can be used to set the teams presence.

        Input:
        [String] availability
        [String] activity
        [Integer] duration_in_m

        Possible combinations:
            - [availability Available, activity Available]
            - [availability Busy, activity InACall]
            - [availability Busy, activity InAConferenceCall]
            - [availability Away, activity Away] 
            - [availability DoNotDisturb, activity Presenting]

        Output:
        [String] response
        """
        if not self.session_uuid:
            raise ValueError("Session uuid needs to be set")
        if duration_in_m < 1 and  duration_in_m > 240:
            raise ValueError("Duration needs to be between 1 and 240")
        if (availability == "Available" and activity == "Available"
           or availability == "Busy" and activity == "InAConferenceCall" 
           or availability == "Busy" and activity == "InACall"
           or availability == "Away" and activity == "Away"
           or availability == "DoNotDisturb" and activity == "Presenting"):
            url = self.base_url+"/endpoints/"
            payload = {
                "id": self.session_uuid,
                "activityReporting": "Transport",
                "deviceType": "Web",
                "availability": availability,
                "activity": activity,
                "expirationDuration": f"PT{duration_in_m}M"
            }
            response = requests.put(url=url, json=payload, headers=self.headers, timeout=self.timeout)
            if response.status_code >=20 and response.status_code <=201:
                return response
            else:
                raise Exception(response)
        else:   
            raise ValueError(f"The Input value {availability}, {activity} is wrong, please use one of the following combinations: \
                    [availability Available, activity Available] \
                    [availability Busy, activity InACall] \
                    [availability Busy, activity InAConferenceCall] \
                    [availability Away, activity Away] \
                    [availability DoNotDisturb, activity Presenting]")

