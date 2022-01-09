# too-many-meetings
Ever thought you spend far too much time in meetings? This powershell script queries your outlook calendar, and gives you the depressing answer you need.

## Usage

First, dot-source the script in to your shell:

1. Open a powershell window and `cd` to the folder you've put the script in.
2. `. ./too-many-meetings.ps1`

Next, you need to get a Microsoft Graph API Token. You can do this easily by:

1. Go to the [MS Graph API Explorer](https://developer.microsoft.com/en-us/graph/graph-explorer)
2. Sign in
3. On the left-hand "Sample queries" pane, scroll down and find the "Outlook Calendar" section.
4. Click on the `[GET] my events for the next week` item and make sure it shows up on the right as `200 - OK - nnnms`
5. Click the `Access token` tab at the top. You should see a massive load of text.
6. Click the icon that looks like two documents to copy it to your clipboard.
7. Go back to powershell, and type `Set-MSGraphAPIToken -Token "TG9yZW0gaXBzdW0gZG9sb3Igc2l0IGFtZXQ..."` where the text in between the quotes is the pasted token you just copied, and press enter.

Now you're ready to run!

1. Change your working hours in the `Invoke-Pain` function, and adjust the value for your daily break.
2. Run the script again by dot-sourcing it: `. ./too-many-meetings.ps1`

## Concept

- My job has a bit of an odd working pattern in the sense that I do technical design meetings (planned work), technical support tickets (mostly unplanned work), and incident response (unplanned work, takes priority). Given the quantity of unplanned work, the planned work would usually suffer. However, if the design meetings are not attended, they plough on without any infosec/secarch guidance whatsoever, and then result in insecure/horribly designed projects and systems. So instead, I attend the meetings as they generate the most business value, the tickets suffer, and I end up being a blocker to progress because people are waiting for things. This script exists to prove with actual figures that we need more people power in my department (and potentially others!) as so much of my time is consumed by meetings. Many of which are low value to the business.
- The script assumes you work from Monday to Friday and have 'some amount' of break each day.
- Hopefully this can help some others manage their time better, or justify to management that you need more resource.
- This was a weekend project :)
