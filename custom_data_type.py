from typing import Union
from pydantic import BaseModel

class Token(BaseModel):
    access_token: str
    token_type: str


class TokenData(BaseModel):
    username: Union[str, None] = None


class User(BaseModel):
    username: str


class UserInDB(User):
    hashed_password: str


class UserInput(BaseModel):
    cylinder1:str=None
    cylinder2:str=None
    cylinder3:str=None
    cylinder4:str=None
    cylinder5:str=None
    cylinder6:str=None
    cylinder7:str=None
    cylinder8:str=None
    cylinder9:str=None
    cylinder10:str=None
    cylinder11:str=None
    cylinder12:str=None
    cylinder13:str=None
    cylinder14:str=None
    cylinder15:str=None

    VESSEL_OBJECT_ID:int=None
    EQUIPMENT_CODE:str=None
    EQUIPMENT_ID:int=None
    JOB_PLAN_ID:int=None
    JOB_ID:int=None
    LOG_ID:int=None

    #Vessel_Information 
    Vessel:str=None
    Hull_No:str=None
    Vessel_Type:str=None
    Local_Start_Time:str=None
    Time_Zone1:str=None
    Local_End_Time:str=None
    Time_Zone2:str=None
    Form_No:str=None
    IMO_No:str=None

    #Engine_info
    Maker:str=None
    Model:str=None
    License_Builder:str=None
    Serial_No:str=None
    MCR:str=None
    Speed_at_MCR:str=None
    Bore:str=None
    Stroke:str=None

    #Turbocharger_info
    Maker_T:str=None
    Model_T:str=None

    #General_Data
    Total_Running_Hour:str=None
    Cylinder_Oil_type:str=None
    Normal_service_load_in_percentage_of_MCR:str=None
    Scrubber:str=None
    Position: str=None
    Cylinder_Oil_feed_rate:str=None
    Inspected_by_Rank:str=None
    Fuel_Sulphur_percentage:str=None
