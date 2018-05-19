# ============================================================================
# DEXTERITY ROBOT TESTS
# ============================================================================
#
# Run this robot test stand-alone:
#
#  $ bin/test -s cshm.content -t test_classroom.robot --all
#
# Run this robot test with robot server (which is faster):
#
# 1) Start robot server:
#
# $ bin/robot-server --reload-path src cshm.content.testing.CSHM_CONTENT_ACCEPTANCE_TESTING
#
# 2) Run robot tests:
#
# $ bin/robot src/plonetraining/testing/tests/robot/test_classroom.robot
#
# See the http://docs.plone.org for further details (search for robot
# framework).
#
# ============================================================================

*** Settings *****************************************************************

Resource  plone/app/robotframework/selenium.robot
Resource  plone/app/robotframework/keywords.robot

Library  Remote  ${PLONE_URL}/RobotRemote

Test Setup  Open test browser
Test Teardown  Close all browsers


*** Test Cases ***************************************************************

Scenario: As a site administrator I can add a Classroom
  Given a logged-in site administrator
    and an add Classroom form
   When I type 'My Classroom' into the title field
    and I submit the form
   Then a Classroom with the title 'My Classroom' has been created

Scenario: As a site administrator I can view a Classroom
  Given a logged-in site administrator
    and a Classroom 'My Classroom'
   When I go to the Classroom view
   Then I can see the Classroom title 'My Classroom'


*** Keywords *****************************************************************

# --- Given ------------------------------------------------------------------

a logged-in site administrator
  Enable autologin as  Site Administrator

an add Classroom form
  Go To  ${PLONE_URL}/++add++Classroom

a Classroom 'My Classroom'
  Create content  type=Classroom  id=my-classroom  title=My Classroom


# --- WHEN -------------------------------------------------------------------

I type '${title}' into the title field
  Input Text  name=form-widgets-IBasic-title  ${title}

I submit the form
  Click Button  Save

I go to the Classroom view
  Go To  ${PLONE_URL}/my-classroom
  Wait until page contains  Site Map


# --- THEN -------------------------------------------------------------------

a Classroom with the title '${title}' has been created
  Wait until page contains  Site Map
  Page should contain  ${title}
  Page should contain  Item created

I can see the Classroom title '${title}'
  Wait until page contains  Site Map
  Page should contain  ${title}
