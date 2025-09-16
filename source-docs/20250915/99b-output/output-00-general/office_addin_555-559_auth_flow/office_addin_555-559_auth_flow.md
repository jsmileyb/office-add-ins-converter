{0}------------------------------------------------

# **Authentication patterns**

07/30/2025

Add-ins may require users to sign-in or sign-up in order to access features and functionality. Input boxes for username and password or buttons that start third party credential flows are common interface controls in authentication experiences. A simple and efficient authentication experience is an important first step to getting users started with your add-in.

## **Best practices**

ノ **Expand table**

| Do                                                                                                                   | Don't                                                                                                     |
|----------------------------------------------------------------------------------------------------------------------|-----------------------------------------------------------------------------------------------------------|
| Prior to sign-in, describe the value of your add-in or<br>demonstrate functionality without requiring an<br>account. | Expect users to sign-in without understanding<br>the value and benefits of your add-in.                   |
| Guide users through authentication flows with a                                                                      | Draw attention to secondary and tertiary tasks                                                            |
| primary, highly visible button on each screen.                                                                       | with competing buttons and calls to action.                                                               |
| Use clear button labels that describe specific tasks<br>like "Sign in" or "Create account".                          | Use vague button labels like "Submit" or "Get<br>started" to guide users through authentication<br>flows. |
| Use a dialog to focus users' attention on                                                                            | Overcrowd your task pane with a first-run                                                                 |
| authentication forms.                                                                                                | experience and authentication forms.                                                                      |
| Find small efficiencies in the flow like auto-focusing                                                               | Add unnecessary steps to the interaction like                                                             |
| on input boxes.                                                                                                      | requiring users to click into form fields.                                                                |
| Provide a way for users to sign out and<br>reauthenticate.                                                           | Force users to uninstall to switch identities.                                                            |

## **Authentication flow**

- 1. First-Run Placemat Place your sign-in button as a clear call-to action inside your addin's first-run experience.

{1}------------------------------------------------

- 2. Identity Provider Choices Dialog Display a clear list of identity providers including a username and password form if applicable. Your add-in UI may be blocked while the authentication dialog is open.

|  | My add-insign-in<br>×<br>My Add-in name<br>Select your sign-in preference<br>Username | My Add-in Name<br>(1) Info - please sign-in with dialog window.<br>× |
|--|---------------------------------------------------------------------------------------|----------------------------------------------------------------------|
|  | Password<br>Sign In<br>Don't have an account? Sign Up<br>Sign in with Microsoft       |                                                                      |
|  |                                                                                       |                                                                      |

- 3. Identity Provider Sign-in The identity provider will have their own UI. Microsoft Entra ID allows customization of sign-in and access panel pages for consistent look and feel with your service. For more information, see [Configure your company branding](https://learn.microsoft.com/en-us/entra/fundamentals/how-to-customize-branding).

{2}------------------------------------------------

| My Add-in Name - Dialog Title | 8                                                                           |                                                                                                                 |
|-------------------------------|-----------------------------------------------------------------------------|-----------------------------------------------------------------------------------------------------------------|
|                               | My Add-in Name                                                              | Converter<br>网<br>(i) Info - Please sign in with the pop-up window.                                             |
|                               | Work or School, or personal Microsoft Account<br>Email or phone<br>Password | Ogo                                                                                                             |
|                               | □ Keep me signed in<br>Sign In                                              | Welcome                                                                                                         |
|                               | Can't access your account?<br>Other sign-in options<br>Get a new account    | Discover what this add-in<br>can do for you today!                                                              |
|                               | 2016 Microsoft<br>  Microsoft<br>Terms of use   Privacy & Cookies           | ള  Achieve more with Office integration<br>Unlock features and functionality<br>Create and visualize like a pro |
|                               |                                                                             | Sign in<br>1                                                                                                    |

- 4. Progress Indicate progress while settings and UI load.

|  | My add-in sign-in<br>×         |                                                  |
|--|--------------------------------|--------------------------------------------------|
|  |                                | My Add-in Name                                   |
|  | My Add-in Brand<br>Authorizing | ① Info - please sign-in with dialog window.<br>× |

### 7 **Note**

When using Microsoft's Identity service you'll have the opportunity to use a branded signin button that is customizable to light and dark themes. Learn more.

## **Single Sign-On authentication flow**

{3}------------------------------------------------

#### 7 **Note**

The single sign-on API is currently supported for Word, Excel, Outlook, and PowerPoint. For more information about single sign-on support, see **[IdentityAPI requirement sets](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/common/identity-api-requirement-sets)**. If you're working with an Outlook add-in, be sure to enable Modern Authentication for the Microsoft 365 tenancy. For information about how to do this, see **[Enable or disable](https://learn.microsoft.com/en-us/exchange/clients-and-mobile-in-exchange-online/enable-or-disable-modern-authentication-in-exchange-online) [modern authentication for Outlook in Exchange Online](https://learn.microsoft.com/en-us/exchange/clients-and-mobile-in-exchange-online/enable-or-disable-modern-authentication-in-exchange-online)**.

Use single sign-on for a smoother end-user experience. The user's identity within Office (either a Microsoft account or a Microsoft 365 identity) is used to sign in to your add-in. As a result users only sign in once. This removes friction in the experience making it easier for your customers to get started.

- 1. As an add-in is being installed, a user will see a consent window similar to the one following:
### 7 **Note**

The add-in publisher will have control over the logo, strings and permission scopes included in the consent window. The UI is pre-configured by Microsoft.

- 2. The add-in will load after the user consents. It can extract and display any necessary user customized information.

{4}------------------------------------------------

### **See also**

- Learn more about developing SSO Add-ins