# ============================================================================
# DEXTERITY ROBOT TESTS
# ============================================================================
#
# Run this robot test stand-alone:
#
#  $ bin/test -s cshm.content -t test_subject.robot --all
#
# Run this robot test with robot server (which is faster):
#
# 1) Start robot server:
#
# $ bin/robot-server --reload-path src cshm.content.testing.CSHM_CONTENT_ACCEPTANCE_TESTING
#
# 2) Run robot tests:
#
# $ bin/robot src/plonetraining/testing/tests/robot/test_subject.robot
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

Scenario: As a site administrator I can add a Subject
  Given a logged-in site administrator
    and an add Subject form
   When I type 'My Subject' into the title field
    and I submit the form
   Then a Subject with the title 'My Subject' has been created

Scenario: As a site administrator I can view a Subject
  Given a logged-in site administrator
    and a Subject 'My Subject'
   When I go to the Subject view
   Then I can see the Subject title 'My Subject'


*** Keywords *****************************************************************

# --- Given ------------------------------------------------------------------

a logged-in site administrator
  Enable autologin as  Site Administrator

an add Subject form
  Go To  ${PLONE_URL}/++add++Subject

a Subject 'My Subject'
  Create content  type=Subject  id=my-subject  title=My Subject


# --- WHEN -------------------------------------------------------------------

I type '${title}' into the title field
  Input Text  name=form-widgets-IBasic-title  ${title}

I submit the form
  Click Button  Save

I go to the Subject view
  Go To  ${PLONE_URL}/my-subject
  Wait until page contains  Site Map


# --- THEN -------------------------------------------------------------------

a Subject with the title '${title}' has been created
  Wait until page contains  Site Map
  Page should contain  ${title}
  Page should contain  Item created

I can see the Subject title '${title}'
  Wait until page contains  Site Map
  Page should contain  ${title}
