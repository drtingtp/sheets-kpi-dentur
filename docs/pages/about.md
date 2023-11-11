---
layout: page
title: About
permalink: /about/
---

## Privacy Policy
- No information is collected or shared by this add-on.
- Using this add-on means granting access to the Google Sheets™ spreadsheet where the add-on is installed ("Datasheet") and allowing the add-on to enhance the usability of the Datasheet.
- To revoke permission that has been granted to this add-on, please refer to [this page]({{ site.baseurl }}/advanced-permission).

### Technically, this add-on
- Installs an installable trigger on the user's behalf for `onChange` events to react to newly added rows in the Datasheet:
    - The `onChange` trigger invokes the function which copies formulas from existing row to the newly added rows in the Datasheet.
- Provides a callable function to clears cells in the Datasheet.