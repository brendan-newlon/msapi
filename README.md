# msapi
An R package for simple, flexible REST API calls to MS Office Graph and SharePoint.

## Installation

	library(devtools)
	install_github("brendan-newlon/msapi")
	library(msapi)
    

## SP

This function allows R to send normally constructed SharePoint queries and return the response as a simple data.frame. 

### ARGUMENTS 
- query
The full URL for the query to send to the MS Office Graph or SharePoint API
- Username 
Your email for authentication on the SharePoint server
- Method 
Defaults to GET but also supports POST. MERGE support hasn't been written in yet.
- acceptLanguage 
Defaults to en
- reenterPassword 
Option to manually clear your password
- refreshTokens 
Option to manually clear all tokens, for example, to connect from a different user account
- tokens_timeout_in_hours 
Set number of hours before new tokens will be requested. The default is 3, but regardless, new token requests will be handled as needed automatically.


### EXAMPLE

	ListItems <- SP("https://{TENANT}.sharepoint.com/sites/_api/Web/Lists(guid'{LIST ID}')/Items", 
	Username = "{USER EMAIL}")

You can use pipes for a clearer view of your queries:

	ListItems <- "https://{TENANT}.sharepoint.com/sites/_api/Web/Lists(guid'{LIST ID}')/Items" %>% 
	SP(Username = "{USER EMAIL}")


For more information about how to construct queries, see the SharePoint REST API documentation: https://docs.microsoft.com/en-us/sharepoint/dev/sp-add-ins/complete-basic-operations-using-sharepoint-rest-endpoints

To further refine your queries, see the OData documentation: https://www.odata.org/documentation/

## MSGraph

This function allows R to send normally constructed MS Office Graph queries and return the response as a simple data.frame. 

For more information about how to construct queries, see the Office Graph Developer Demo: https://developer.microsoft.com/en-us/graph/graph-explorer

To further refine your queries, see the OData documentation: https://www.odata.org/documentation/


### ARGUMENTS 

- query 
The full URL for the query to send to the MS Office Graph or SharePoint API
- handle_pagination 
Defaults to TRUE to automatically continue requesting until all data has been received
- reset_token 
Option to manually reset your access token
- assign_responses 
Option to also assign the response to your global environment as the variable MSGraph_latestResponse. This is helpful for debugging if your response isn't what you expected, or if you want to have the option to interrupt the operation while keeping as much data as has been received up to that point.

### EXAMPLE:
  
	User_IDs <- MSGraph("https://graph.microsoft.com/v1.0/users?$select=id")

	User_IDs <- "https://graph.microsoft.com/v1.0/users?$select=id" %>% MSGraph()
