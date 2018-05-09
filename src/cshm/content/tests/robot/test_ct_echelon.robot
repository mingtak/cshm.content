# ============================================================================
# DEXTERITY ROBOT TESTS
# ============================================================================
#
# Run this robot test stand-alone:
#
#  $ bin/test -s cshm.content -t test_echelon.robot --all
#
# Run this robot test with robot server (which is faster):
#
# 1) Start robot server:
#
# $ bin/robot-server --reload-path src cshm.content.testing.CSHM_CONTENT_ACCEPTANCE_TESTING
#
# 2) Run robot tests:
#
# $ bin/robot src/plonetraining/testing/tests/robot/test_echelon.robot
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

Scenario: As a site administrator I can add a Echelon
  Given a logged-in site administrator
    and an add Echelon form
   When I type 'My Echelon' into the title field
    and I submit the form
   Then a Echelon with the title 'My Echelon' has been created

Scenario: As a site administrator I can view a Echelon
  Given a logged-in site administrator
    and a Echelon 'My Echelon'
   When I go to the Echelon view
   Then I can see the Echelon title 'My Echelon'


*** Keywords *****************************************************************

# --- Given ------------------------------------------------------------------

a logged-in site administrator
  Enable autologin as  Site Administrator

an add Echelon form
  Go To  ${PLONE_URL}/++add++Echelon

a Echelon 'My Echelon'
  Create content  type=Echelon  id=my-echelon  title=My Echelon


# --- WHEN -------------------------------------------------------------------

I type '${title}' into the title field
  Input Text  name=form-widgets-IBasic-title  ${title}

I submit the form
  Click Button  Save

I go to the Echelon view
  Go To  ${PLONE_URL}/my-echelon
  Wait until page contains  Site Map


# --- THEN -------------------------------------------------------------------

a Echelon with the title '${title}' has been created
  Wait until page contains  Site Map
  Page should contain  ${title}
  Page should contain  Item created

I can see the Echelon title '${title}'
  Wait until page contains  Site Map
  Page should contain  ${title}
