import pynput.keyboard
import time


USER_ID_POSITION = (550, 474)
PASSWORD_POSITION = (550, 586)
LOGIN_POSITION = (770, 692)

ANIMAL_HEALTH_POSITION = (75, 457)
VACCINATION_POSITION = (75, 498)
WITHOUT_CAMPAIGN_POSITION = (371, 298)
BATCH_NUMBER_POSITION = (290, 505)

mouse = pynput.mouse.Controller()
# time.sleep(2)
mouse.position = BATCH_NUMBER_POSITION
print(mouse.position)
