# constants.py

import os

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

DATA_DIR = os.path.join(BASE_DIR, "data")

PREBOOK_FILE = os.path.join(
    DATA_DIR,
    "prebook_schedule.xlsx"
)

FIXED_FILE = os.path.join(
    DATA_DIR,
    "fixed_schedule.xlsx"
)


TIME_OPTIONS = [
    "6:00am", "6:30am", "7:00am", "7:30am",
    "8:00am", "8:30am", "9:00am", "9:30am",
    "10:00am", "10:30am", "11:00am", "11:30am",
    "12:00pm", "12:30pm", "1:00pm", "1:30pm",
    "2:00pm", "2:30pm", "3:00pm", "3:30pm",
    "4:00pm", "4:30pm", "5:00pm", "5:30pm", "6:00pm"
]

ROLE_OPTIONS = [
    "值班",
    "卫生",
]

ROLES = ["值班", "卫生", "佛台", "供台", "供花", "供果", "膳食", "佛学班"]

ROLE_TEXT = {
    "值班": {"zh": "值班", "en": "Duty"},
    "卫生": {"zh": "卫生", "en": "Cleaning"},
    "佛台": {"zh": "佛台", "en": "Altar"},
    "供台": {"zh": "供台", "en": "Offering Table"},
    "供花": {"zh": "供花", "en": "Flowers"},
    "供果": {"zh": "供果", "en": "Fruit Offering"},
    "膳食": {"zh": "膳食组", "en": "Meal Team"},
    "佛学班": {"zh": "佛学班", "en": "Buddhist Class"},
}

ROLE_DEFAULT_TIME = {
    "卫生": ("08:00", "10:00"),
    "佛台": ("08:00", "10:00"),
    "供台": ("06:00", "08:00"),
    "供花": ("08:00", "10:00"),
    "供果": ("08:00", "10:00"),
    "膳食": ("08:00", "14:00"),
    "佛学班": ("08:00", "12:00"),
}

MEAL_MAX_SIGNUP = 9