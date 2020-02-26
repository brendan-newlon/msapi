#' SP
#'
#' @description
#' Simple, flexible REST API calls to MS SharePoint.
#'
#' @details
#' This function allows R to send normally constructed SharePoint queries
#' and return the response as a simple data.frame.
#'
#' For more information about how to construct queries, see the SharePoint REST API documentation:
#' https://docs.microsoft.com/en-us/sharepoint/dev/sp-add-ins/complete-basic-operations-using-sharepoint-rest-endpoints
#'
#' To further refine your queries, see the OData documentation:
#' https://www.odata.org/documentation/
#'
#'
#' @param query The full URL for the query to send to the MS Office Graph or SharePoint API
#' @param Username Your email for authentication on the SharePoint server
#' @param Method Defaults to GET but also supports POST. MERGE support hasn't been written in yet.
#' @param acceptLanguage Defaults to en
#' @param reenterPassword Option to manually clear your password
#' @param refreshTokens Option to manually clear all tokens, for example, to connect from a different user account
#' @param tokens_timeout_in_hours Set number of hours before new tokens will be requested.
#' The default is 3, but regardless, new token requests will be handled as needed automatically.
#'
#'
#' @return A data.frame with the complete, pagerized response to your query.
#'
#'
#' @examples
#' ListItems <- SP("https://{YOUR TENANT}.sharepoint.com/sites/_api/Web/Lists(guid'{THE ID OF YOUR LIST}')/Items", Username = "{YOUR AUTHENTICATION EMAIL}")
#'
#'
#' You can use pipes for a clearer view of your queries:
#'
#' ListItems <- "https://{YOUR TENANT}.sharepoint.com/sites/_api/Web/Lists(guid'{THE ID OF YOUR LIST}')/Items" %>% SP(Username = "{YOUR AUTHENTICATION EMAIL}")

SP <-
  function(query = "",
           Username = "",
           Method = "GET",
           acceptLanguage = "en",
           reenterPassword = FALSE,
           refreshTokens = FALSE,
           tokens_timeout_in_hours = "3",
           ...) {

    if (!exists("Username") || Username == "" || is_empty(Username)) {
      Username <- rstudioapi::askForSecret("Enter your SharePoint Username. eg. me@mycompany.com")}

    # They STILL didn't enter a username...
    if (!exists("Username") ||Username == "" || is_empty(Username)) { stop("Try again with a username.")}

    if (query == "") {
      query <- dlgInput("Enter your query, eg. https://cclonline.sharepoint.com/_api/web/ ")$res %>%
        gsub(" ","%20",.) # to accomodate spaces in queries in R, but send them properly to web
    }

    if (reenterPassword == "TRUE") {if ("SP_p" %in% key_list("SP_p")$service) {key_delete("SP_p")}}
    if (refreshTokens == "TRUE") {sp_clear_keys()}

    # if more hours passed than tokens_timeout_in_hours argument (default = 3), then get new tokens
    if ("SP_connection_time" %in% key_list("SP_connection_time")$service) {
      SP_last_connection <- as.POSIXct(key_get("SP_connection_time"))
      if (Sys.time() > SP_last_connection + hours(tokens_timeout_in_hours)) {
        sp_clear_keys()
      }
    } else {
      t <- as.character(Sys.time())
      key_set_with_value("SP_connection_time", password = t)
    }

    # Check if it's paginated, if so, get all pages of results and bind them together
    response <- SP.handle.pagination(Method, query, Username)

    return(response)
  } # end of function SP()


#' MSGraph
#'
#' @description
#' Simple, flexible REST API calls to MS Office Graph
#'
#' @details
#' This function allows R to send normally constructed MS Office Graph queries
#' and return the response as a simple data.frame.
#'
#' For more information about how to construct queries, see the Office Graph Developer Demo:
#' https://developer.microsoft.com/en-us/graph/graph-explorer
#'
#' To further refine your queries, see the OData documentation:
#' https://www.odata.org/documentation/
#'
#'
#'
#' @param query The full URL for the query to send to the MS Office Graph or SharePoint API
#' @param handle_pagination Defaults to TRUE to automatically continue requesting until all data has been received
#' @param reset_token Option to manually reset your access token
#' @param assign_responses Option to also assign the response to your global environment
#' as the variable MSGraph_latestResponse. This is helpful for debugging if your response
#' is not what you expected, or if you want to have the option to interrupt the operation
#' while keeping as much data as has been received up to that point.
#'
#'
#' @return A data.frame with the complete, pagerized response to your query.
#'
#'
#' @examples
#' User_IDs <- MSGraph("https://graph.microsoft.com/v1.0/users?$select=id")
#'
#'
#' You can use pipes for a clearer view of your queries:
#'
#' User_IDs <- "https://graph.microsoft.com/v1.0/users?$select=id" %>% MSGraph()
#'

#__________________________________________________MSGraph()  ########## user-facing function
MSGraph <- function(query ="", handle_pagination = TRUE, reset_token = FALSE, assign_responses = FALSE) {
  if (query == "") {query <- dlgInput("Enter your MS Graph query:")$res } # Prompt for query if none entered
  if (reset_token == TRUE){ MSGraph_delete_token()} # option to delete token
  response3 <- MSGraph_handle_pagination(query = query, handle_pagination = handle_pagination, assign_responses = assign_responses)
  response3 <-  response3[!str_detect(response3$name,"@odata\\."),]
  return(response3)
} ############### End of function: MSGraph() #################




# Dependencies:
ipak <- function(pkg){
  new.pkg <- pkg[!(pkg %in% installed.packages()[, "Package"])]
  if (length(new.pkg))
    install.packages(new.pkg, dependencies = TRUE)
  sapply(pkg, require, character.only = TRUE)}

packages <- c("httr","jsonlite", "utf8","curl", "lubridate", "tidyr", "keyring", "dplyr", "purrr", "rlang", "xml2", "svDialogs","stringr","magrittr")
ipak(packages)

as.df <- function(x) {x <- as.data.frame(x,stringsAsFactors = FALSE)}
`%notin%` <- Negate(`%in%`)



#####################################################################
################## MS SharePoint Functions ##########################
#####################################################################
#
#++++++++++++++++++++  ABOUT  +++++++++++++++++++
# This set of functions is daisy-chained to isolate steps and make debugging easier.
#
# The user only needs to interact with the function SP()
# The main function passes work down the chain: SP() --> SP.handle.pagination() --> SP.send.query()
# Other functions below support them to establish authentication, clean the returned data, etc.
#================================================

## ____________________________________________________________________ Function getSP_SecToken()

getSP_SecToken <- function(query, Username) {
  if (!exists("Address")) {
    Address = paste0("https://", regmatches(query,regexpr("[[:alnum:]]{1,}\\.sharepoint\\.com", query)))
  }
  Address_base = regmatches(Address,regexpr("[[:alnum:]]{1,}\\.sharepoint\\.com", Address)) # remove https:// from address

   # request = suppressWarnings(readLines("saml.xml")) # read XML soap envelope
  request <- "<s:Envelope xmlns:s=\"http://www.w3.org/2003/05/soap-envelope\"       xmlns:a=\"http://www.w3.org/2005/08/addressing\"       xmlns:u=\"http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-utility-1.0.xsd\">   <s:Header>     <a:Action s:mustUnderstand=\"1\">http://schemas.xmlsoap.org/ws/2005/02/trust/RST/Issue</a:Action>     <a:ReplyTo>       <a:Address>http://www.w3.org/2005/08/addressing/anonymous</a:Address>     </a:ReplyTo>     <a:To s:mustUnderstand=\"1\">https://login.microsoftonline.com/extSTS.srf</a:To>     <o:Security s:mustUnderstand=\"1\"        xmlns:o=\"http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-secext-1.0.xsd\">       <o:UsernameToken>         <o:Username>{Username}</o:Username>         <o:Password>{Password}</o:Password>       </o:UsernameToken>     </o:Security>   </s:Header>   <s:Body>     <t:RequestSecurityToken xmlns:t=\"http://schemas.xmlsoap.org/ws/2005/02/trust\">       <wsp:AppliesTo xmlns:wsp=\"http://schemas.xmlsoap.org/ws/2004/09/policy\">         <a:EndpointReference>           <a:Address>{Address}</a:Address>         </a:EndpointReference>       </wsp:AppliesTo>       <t:KeyType>http://schemas.xmlsoap.org/ws/2005/05/identity/NoProofKey</t:KeyType>       <t:RequestType>http://schemas.xmlsoap.org/ws/2005/02/trust/Issue</t:RequestType>       <t:TokenType>urn:oasis:names:tc:SAML:1.0:assertion</t:TokenType>     </t:RequestSecurityToken>   </s:Body> </s:Envelope>"
  request = gsub("\\{Username\\}", Username, request) # paste username into XML form
  if ("SP_p" %in% key_list("SP_p")$service) {
    SP_pass <- key_get("SP_p")
  } else {
    SP_pass <- rstudioapi::askForSecret("Enter SharePoint password")
    key_set_with_value("SP_p", password = SP_pass)
  }
  request = gsub("\\{Password\\}", SP_pass, request) # paste password into XML form

  request = gsub("\\{Address\\}", Address_base, request) # paste address into XML form
  response = httr::POST(url = "https://login.microsoftonline.com/extSTS.srf", body = request) # request security token from microsoft online
  if (response$status_code != 200)
    stop("Receiving security token failed.")# Check if request was successful
  content = as_list(read_xml(rawToChar(response$content))) # decode response content
  SecToken = as.character(
    content$Envelope$Body$RequestSecurityTokenResponse$RequestedSecurityToken$BinarySecurityToken
  ) # extract security token

  # if that fails, and SecToken is empty, extract it this way:
  if ("SP_SecToken" %in% key_list("SP_SecToken")$service) {
    SecToken <- key_get("SP_SecToken")
  }

  if(is_empty(SecToken)){
    content <- read_xml(rawToChar(response$content))
    SecToken <- as.character(xml_child(xml_child(xml_child(xml_child(content, 2), 1), 4), 1)) # this is analogous to here:
    #content$Body$RequestSecurityTokenResponse$RequestedSecurityToken$BinarySecurityToken
  }

  # try again if password bad?
  if (is_empty(SecToken)) {
    SP_pass <-
      rstudioapi::askForSecret(
        "No security token was received. The password you entered may have been incorrect. Try again"
      )
    key_set_with_value("SP_p", password = SP_pass)
    getSP_SecToken() # is it odd to tell the function to start itself over?
  } else {
    key_set_with_value("SP_SecToken", password = SecToken)
  }
  assign("SecToken", SecToken, envir = parent.frame())
} # end of function


# ____________________________________________________________________ Function getSP_AuthToken()
getSP_AuthToken <- function(query, Username) {
  if ("SP_SecToken" %in% key_list("SP_SecToken")$service) {
    SecToken <- key_get("SP_SecToken")
  } else {
    getSP_SecToken(query, Username)
  }

  if (!exists("Address")) {
    Address = paste0("https://", regmatches(query,regexpr("[[:alnum:]]{1,}\\.sharepoint\\.com", query)))
  }
  Address_base = regmatches(Address,regexpr("[[:alnum:]]{1,}\\.sharepoint\\.com", Address)) # remove https:// from address

  # Step 2: Use security token to request Access token
  response = httr::POST(paste0("https://", Address_base, "/_forms/default.aspx?wa=wsignin1.0"),body = SecToken,add_headers(Host = Address_base, Accept = "application/json;odata=verbose")
  ) # post security token to sharepoint online

  #### Here you need from the response$headers$'set cookie'   -- the ones beginning with rtFa=   and FedAuth=
  cookie = paste0("rtFa=",
                  response$cookies$value[response$cookies$name %in% "rtFa"],
                  "; FedAuth=",
                  response$cookies$value[response$cookies$name %in% "FedAuth"]) # concatenate cookies for header
  # token <- cookie
  key_set_with_value("SP_token", password = cookie)
  # key_get("SP_token")
  auth <- paste0("Bearer ", cookie)

    assign("auth",auth, envir = parent.frame())

  key_set_with_value("SP_auth", password = auth)
}

# ____________________________________________________________________ Function sp_clear_keys()

sp_clear_keys <- function() {
  if ("SP_token" %in% key_list("SP_token")$service) {key_delete("SP_token")}
  if ("SP_p" %in% key_list("SP_p")$service) {key_delete("SP_p")}
  if ("SP_SecToken" %in% key_list("SP_SecToken")$service) {key_delete("SP_SecToken")}
  if ("SP_auth" %in% key_list("SP_auth")$service) {key_delete("SP_auth")}
  if ("SP_connection_time" %in% key_list("SP_connection_time")$service) {key_delete("SP_connection_time")}
  # rm(SPCon,SP_pass,SP_query,SP_response,SP_response_clean,SP_response_JSON,SP_response_raw,SecToken,token,response_JSON,response_raw,query,headers,Address_base,Address,acceptLanguage,auth,cookie,request,Username,getSP_SecToken,response,response1,response2,response3,response_clean,content)
}


# ____________________________________________________________________ Function sp_cleanup_variables()

sp_cleanup_variables <- function() {
  sp_clear_keys()
  rm(SPCon,SP_pass,SP_query,SP_response,SP_response_clean,SP_response_JSON,SP_response_raw,SecToken,token,response_JSON,response_raw,query,headers,Address_base,Address,acceptLanguage,auth,cookie,request,Username,getSP_SecToken,response,response1,response2,response3,response_clean,content)
}


# ____________________________________________________________________ Function SP.send.query()

SP.send.query <- function(Method, query, Username) {
  ##################################### Here we actually send the query
  if ("SP_auth" %in% key_list("SP_auth")$service) {
    auth <- key_get("SP_auth")
  } else {
    getSP_AuthToken(query, Username)
  }
    Address = paste0("https://", regmatches(query,regexpr("[[:alnum:]]{1,}\\.sharepoint\\.com", query)))
  Address_base = regmatches(Address,regexpr("[[:alnum:]]{1,}\\.sharepoint\\.com", Address)) # remove https:// from address

  if(Method == "GET"){ SPCon <- GET(query,add_headers(Authorization = auth,Host = Address_base,Accept = "application/json;odata=verbose"))
  } else {
    if(Method == "POST"){SPCon <-POST(query,add_headers(Authorization = auth,Host = Address_base,Accept = "application/json;odata=verbose"))
    # Not sure how to handle MERGE requests
    }}

  # This is the returned data:
  response_JSON <- content(SPCon, "text")
  # save the original JSON data response - helps as a reference when cleaning breaks something
  # response <- fromJSON(response_JSON)
  # response_orig <- response

  ### -Access token is expired or bad?
  if (grepl("UnauthorizedAccessException", response_JSON) ||
      grepl("Unsupported .* token", response_JSON)        ) {
    response <- fromJSON(response_JSON)
    sp_cleanup_variables()
    # Try again with different credentials
    stop(paste0("The connection was refused with an error:\n\n",response$error$code,"\n",response$error$message$value,"\n\n","Try with different credentials or refresh your tokens using the argument refreshTokens = TRUE.\nIf that fails, consult your SharePoint administrator or tech support staff for more help."))
  }

  # assign("SP_response_JSON", response_JSON, envir = .GlobalEnv)
  return (response_JSON)
} # end function


# ____________________________________________________________________ Function SP.handle.pagination()  & supporting functions

recursive.unlist.cols <- function(l) {
  if (!is.list(l)) return(l)
  do.call('rbind', lapply(l, function(x) `length<-`(x, max(lengths(l)))))
}

super.flatten <- function(x){x <- x %>% jsonlite::flatten(recursive = TRUE) %>% recursive.unlist.cols() %>% t() %>% as.df() %>% mutate_all(as.character)}


remove.empty.cols <- function(df){
  df <- df[,colSums(is.na(df))<nrow(df)]
  df <- Filter(function(x) !(all(x==""|x==0|x=="NULL"|x=="NA")), df)
  return(df)
}

cleaner.names <- function(x) {
  names(x) <- names(x) %>% gsub("_x0020_|\\.","_",.)
  names(x) <- names(x) %>% gsub("__|___","_",.)
  names(x) <- names(x) %>% gsub("^_|_x0027_","",.)
  return(x)
}


SP.handle.pagination <- function(Method, query, Username){
  response_JSON <-  SP.send.query(Method, query, Username)
  response <- fromJSON(response_JSON)

  response_prev <- fromJSON(response_JSON)$d$results %>% super.flatten()

  # Set page count for update in the console
  i <- 2
  repeat {
    #_______________________ if pagination...
    if ("__next" %notin% names(response$d)) {
      return(response_prev)
      break # there's no (MORE) pagination
    } else {
      msg <- paste0("Getting paginated data, page=", i)
      print(eval(msg, envir = globalenv()))
      i <- i + 1
      query <- response$d$`__next`
      response_JSON <-  SP.send.query(Method, query, Username)
      response <- fromJSON(response_JSON)

      response_next <- fromJSON(response_JSON)$d$results %>% super.flatten()


      # Join new results to previous
      response_prev <- bind_rows(response_prev,response_next)
    } # end else (handling paginated data)
  } # end repeat
  response_prev <- response_prev %>% remove.empty.cols() %>% cleaner.names()
  return(response_prev)
} # end function


# ____________________________________________________________________ Function SP()
# User-facing function
# SP <-
#   function(query = "",
#            Username = "",
#            Method = "GET",
#            acceptLanguage = "en",
#            reenterPassword = FALSE,
#            refreshTokens = FALSE,
#            tokens_timeout_in_hours = "3",
#            ...) {
#
#     if (!exists("Username") || Username == "" || is_empty(Username)) {
#       Username <- rstudioapi::askForSecret("Enter your SharePoint Username. eg. me@mycompany.com")}
#
#     # They STILL didn't enter a username...
#     if (!exists("Username") ||Username == "" || is_empty(Username)) { stop("Try again with a username.")}
#
#     if (query == "") {
#       query <- dlgInput("Enter your query, eg. https://cclonline.sharepoint.com/_api/web/ ")$res %>%
#         gsub(" ","%20",.) # to accomodate spaces in queries in R, but send them properly to web
#     }
#
#     if (reenterPassword == "TRUE") {if ("SP_p" %in% key_list("SP_p")$service) {key_delete("SP_p")}}
#     if (refreshTokens == "TRUE") {sp_clear_keys()}
#
#     # if more hours passed than tokens_timeout_in_hours argument (default = 3), then get new tokens
#     if ("SP_connection_time" %in% key_list("SP_connection_time")$service) {
#       SP_last_connection <- as.POSIXct(key_get("SP_connection_time"))
#       if (Sys.time() > SP_last_connection + hours(tokens_timeout_in_hours)) {
#         sp_clear_keys()
#       }
#     } else {
#       t <- as.character(Sys.time())
#       key_set_with_value("SP_connection_time", password = t)
#     }
#
#     # Check if it's paginated, if so, get all pages of results and bind them together
#     response <- SP.handle.pagination(Method, query, Username)
#
#     return(response)
#   } # end of function SP()


#####################################################################
################ MS Office Graph Functions ##########################
#####################################################################
#
#++++++++++++++++++++  ABOUT  +++++++++++++++++++
# This set of functions is daisy-chained to isolate steps and make debugging easier.
#
# The user only needs to interact with the function MSGraph()
# The main function passes work down the chain: MSGraph() --> MSGraph_handle_pagination() --> MSGraph_clean_response() --> MSGraph_send_query
# Other functions below support them to establish authentication, clean the returned data, etc.
#================================================

#
# MSGraph() and related functions
# MSGraph(query, handle_pagination = TRUE, reset_token = FALSE, assign_responses = FALSE)
#
# eg:  to get the IDs of all users in the account:
#   users <- MSGraph("https://graph.microsoft.com/v1.0/users?$select=id", assign_responses = TRUE)

# Note: The functions below are daisy-chained to simplify debugging; MSGraph() is the only user-facing function.
# Note: This currently works with MS Graph developer preview. It can be adapted for app & simpler authentication like SP() above.

# library(httr)
# library(keyring)
# library(jsonlite)
# library(magrittr)
# library(stringr)
# library(svDialogs)
# as.df <- function(x) {x <- as.data.frame(x,stringsAsFactors = FALSE)}
# `%notin%` <- Negate(`%in%`)

#__________________________________________________MSGraph_delete_token()
MSGraph_delete_token <- function(){if ("MSO_Graph_Token" %in% key_list("MSO_Graph_Token")$service) {key_delete("MSO_Graph_Token")}}

#__________________________________________________MSGraph_handle_auth()
MSGraph_handle_auth <- function(){
  if ("MSO_Graph_Token" %in% key_list("MSO_Graph_Token")$service) {
    auth <- key_get("MSO_Graph_Token")
  } else {
    # Get new auth
    browseURL("https://developer.microsoft.com/en-us/graph/graph-explorer/preview",browser = getOption("browser"),encodeIfNeeded = FALSE)
    token <- rstudioapi::askForSecret(
      "Enter your authentication token.\n
    1. Sign in at https://developer.microsoft.com/en-us/graph/graph-explorer/preview \n
    2. Click \"Auth\" then copy & paste your token here:"
    )
    auth <- paste0("Bearer ", token)
    key_set_with_value("MSO_Graph_Token", password = auth)
  }
  return(auth)
}

#__________________________________________________MSGraph_send_query()
MSGraph_send_query <- function(query){
  auth <- MSGraph_handle_auth()
  GraphCon <- GET(query, add_headers(Authorization = auth, Host = "graph.microsoft.com"))
  response <- content(GraphCon, "text")   # This is the returned data as JSON
  return(response)
}

#__________________________________________________MSGraph_clean_response()
MSGraph_clean_response <- function(query) {
  response <- MSGraph_send_query(query)
  # save the original JSON data response - helps as a reference for debugging bad responses
  response_JSON <- response
  response1 <- fromJSON(response)

  response2 <- suppressWarnings(response1 %>%
                                  unlist() %>%
                                  enframe %>%
                                  unnest)

  response3 <- as.df(response2)   # Flatten to dataframe

  # Handle Invalid Authentication response
  if (isTRUE(grepl("InvalidAuthentication", response3$value[1]))) {
    MSGraph_delete_token()
    auth <- MSGraph_handle_auth()
    if (isTRUE(grepl("InvalidAuthentication", response3$value[1]))) {stop("The token is invalid. Try again with a valid token.")}
  }
  return(response3)
}

#__________________________________________________MSGraph_handle_pagination()
MSGraph_handle_pagination <- function(query, handle_pagination, assign_responses){
  response3 <- suppressWarnings(MSGraph_clean_response(query))

  if(handle_pagination){
    i <- 2
    repeat {
      if ("TRUE" %in% str_detect(response3$name[2], "@odata.nextLink")) {
        msg <- paste0("Getting paginated data, page=", i)
        print(eval(msg, envir = globalenv()))

        i <- i + 1
        nextlink <- response3[str_detect(response3$name, "@odata.nextLink"), 2]
        response3_cont <- suppressWarnings(MSGraph_clean_response(query = nextlink))
        response3 <-  response3[!str_detect(response3$name,"@odata\\."),]   # remove any lines that have "@odata."
        response3 <- rbind(response3_cont, response3)

        if(assign_responses){
          assign("MSGraph_latestResponse", response3, envir = .GlobalEnv)
          # better if this could assign a list of ALL responeses, raw and cleaned !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
        }

      } else {
        if(assign_responses){
          assign("MSGraph_latestResponse", response3, envir = .GlobalEnv)
          # better if this could assign a list of ALL responeses, raw and cleaned !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
        }

        break
      }
    }
  }
  return(response3)
}

# #__________________________________________________MSGraph()  ########## user-facing function
# MSGraph <- function(query ="", handle_pagination = TRUE, reset_token = FALSE, assign_responses = FALSE) {
#   if (query == "") {query <- dlgInput("Enter your MS Graph query:")$res } # Prompt for query if none entered
#   if (reset_token == TRUE){ MSGraph_delete_token()} # option to delete token
#   response3 <- MSGraph_handle_pagination(query = query, handle_pagination = handle_pagination, assign_responses = assign_responses)
#   response3 <-  response3[!str_detect(response3$name,"@odata\\."),]
#   return(response3)
# } ############### End of function: MSGraph() #################

