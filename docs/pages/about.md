---
layout: page
title: About
permalink: /about/
---

# Author

This add-on is written by [drtingtp](https://github.com/drtingtp) for the FT Kuala Lumpur and Putrajaya State Oral Health Division, Ministry of Health, Malaysia.

# Privacy Policy

- No information is collected or shared by this add-on.
- You ("User") using this add-on means granting access to the Google Sheets™ spreadsheet where the add-on is installed ("Datasheet") and allowing the add-on to enhance the usability of the Datasheet.
- To revoke permission that has been granted to this add-on, please refer to [this page]({{ site.baseurl }}/advanced/review-permission).

Technically, this add-on requires the User's permission to
- Install an installable trigger on the User's behalf for `onChange` events to react to newly added rows in the Datasheet:
    - The `onChange` trigger invokes the function which copies formulas from existing row to the newly added rows in the Datasheet.
- Provide a callable function to clears cells in the Datasheet.