# TECHNICAL INFO

This web part is intended to be used on SharePoint Online modern pages.

HTML content in this web part is displayed in an IFRAME.<br/>
This means that this web part should work fine of SharePoint sites that have "custom scripts" set to "blocked".

To improve the security for pages that use this web part, the IFRAME is sandboxed.<br/>
IFRAME HTML content can use scripts and can redirect top frame on user activation:
- allow-scripts
- allow-top-navigation-by-user-activation

**Other functionality within IFRAME is not allowed.**<br/>

As a result, for example, embedded script's access to user's cookies is not possible within the IFRAME.<br/>
Consequently, if you want to embed external content that for some reason wants access to user's cookies, such components won't work within this web part.