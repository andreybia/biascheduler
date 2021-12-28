Imports System.Text
Imports BIAScheduler.SchedulerFunctions
Imports RestSharp
Imports BIAScheduler.DBConnector


Public Class SeonAPI

    Public Shared Function CreateRequestBody(checkType As String, useID As String, ipFrom As String, emailFrom As String, phoneFrom As String) As String 'check type : 1-ip, 2-email, 3-phone number
        Dim ipForCheck As String = ""
        Dim emailForCheck As String = ""
        Dim phoneForCheck As Double = 0

        Select Case checkType
            Case 1
                ipForCheck = ipFrom
            Case 2
                emailForCheck = emailFrom
            Case 3
                phoneForCheck = phoneFrom
        End Select



        Dim returnValue As String = "{" & vbLf & "  ""config"": {" & vbLf & "    ""ip"": {" & vbLf &
            "      ""include"": ""flags,history,id""," & vbLf & "      ""timeout"": 2000," & vbLf &
            "      ""version"": ""v1.0""" & vbLf & "    }," & vbLf & "    ""email"": {" & vbLf &
            "      ""include"": ""flags,history,id""," & vbLf & "      ""timeout"": 2000," & vbLf &
            "      ""version"": ""v2.0""" & vbLf & "    }," & vbLf & "    ""phone"": {" & vbLf &
            "      ""include"": ""flags,history,id""," & vbLf & "      ""timeout"": 2000," & vbLf &
            "      ""version"": ""v1.0""" & vbLf & "    }," & vbLf & "    ""ip_api"": true," & vbLf &
            "    ""email_api"": true," & vbLf & "    ""phone_api"": true," & vbLf &
            "    ""device_fingerprinting"": true," & vbLf &
            "    ""ignore_velocity_rules"": false," & vbLf & "    ""response_fields"": ""id,state,fraud_score,ip_details,email_details,phone_details,bin_details,version,applied_rules,device_details,calculation_time,seon_id""" & vbLf & "  }," & vbLf &
            "	    ""ip"": """ & ipForCheck & """," & vbLf & "        ""action_type"": """"," & vbLf &
            "        ""transaction_id"": """"," & vbLf & "        ""affiliate_id"": """"," & vbLf &
            "        ""affiliate_name"": """"," & vbLf & "        ""order_memo"": """"," & vbLf &
            "        ""email"": """ & emailForCheck & """," & vbLf & "        ""email_domain"": """"," & vbLf &
            "        ""password_hash"": """"," & vbLf & "        ""user_fullname"": """"," & vbLf &
            "        ""user_name"": """"," & vbLf & "        ""user_id"": """ & useID & """," & vbLf &
            "        ""user_dob"": """"," & vbLf & "        ""user_category"": """"," & vbLf &
            "        ""user_account_status"": """"," & vbLf & "        ""user_created"": """"," & vbLf &
            "        ""user_country"": """"," & vbLf & "        ""user_city"": """"," & vbLf &
            "        ""user_region"": """"," & vbLf & "        ""user_zip"": """"," & vbLf &
            "        ""user_street"": """"," & vbLf & "        ""user_street2"": """"," & vbLf &
            "        ""device_id"": """"," & vbLf & "        ""session"": """"," & vbLf &
            "        ""payment_mode"": """"," & vbLf & "        ""card_fullname"": """"," & vbLf &
            "        ""card_bin"": """"," & vbLf & "        ""card_hash"": """"," & vbLf &
            "        ""card_last"": """"," & vbLf & "        ""card_expire"": """"," & vbLf &
            "        ""avs_result"": """"," & vbLf & "        ""cvv_result"": """"," & vbLf &
            "        ""receiver_fullname"": """"," & vbLf & "        ""receiver_bank_account"": """"," & vbLf &
         "        ""sca_method"": """"," & vbLf & "        ""user_bank_account"": """"," & vbLf &
        "        ""user_bank_name"": """"," & vbLf & "        ""user_balance"": """"," & vbLf &
        "        ""user_verification_level"": """"," & vbLf & "        ""status_3d"": """"," & vbLf &
        "        ""regulation"": """"," & vbLf & "        ""payment_provider"": """"," & vbLf &
        "        ""phone_number"": " & phoneForCheck & "," & vbLf & "        ""transaction_type"": """"," & vbLf &
        "        ""transaction_amount"": """"," & vbLf & "        ""transaction_currency"": """"," & vbLf &
        "        ""brand_id"": """"," & vbLf & "        ""items"": [{" & vbLf & "        	""item_id"": """"," & vbLf &
        "        	""item_quantity"": """"," & vbLf & "        	""item_name"": """"," & vbLf &
        "        	""item_price"": """"," & vbLf & "        	""item_store"": """"," & vbLf &
        "        	""item_store_country"": """"," & vbLf & "        	""item_category"": """"," & vbLf &
        "        	""item_url"": """"," & vbLf & "        	""item_custom_fields"": {}" & vbLf &
        "         }, {" & vbLf & "          ""item_id"": """"," & vbLf & "        	""item_quantity"": """"," & vbLf &
        "        	""item_name"": """"," & vbLf & "        	""item_price"": """"," & vbLf &
        "        	""item_store"": """"," & vbLf & "        	""item_store_country"": """"," & vbLf &
        "        	""item_category"": """"," & vbLf & "        	""item_url"": """"," & vbLf &
        "        	""item_custom_fields"": {}" & vbLf & "	}]," & vbLf &
        "        ""shipping_country"": """"," & vbLf & "        ""shipping_city"": """"," & vbLf &
        "        ""shipping_region"": """"," & vbLf & "        ""shipping_zip"": """"," & vbLf &
        "        ""shipping_street"": """"," & vbLf & "        ""shipping_street2"": """"," & vbLf &
        "        ""shipping_phone"": """"," & vbLf & "        ""shipping_fullname"": """"," & vbLf &
        "        ""shipping_method"": """"," & vbLf & "        ""billing_country"": """"," & vbLf &
        "        ""billing_city"": """"," & vbLf & "        ""billing_region"": """"," & vbLf &
        "        ""billing_zip"": """"," & vbLf & "        ""billing_street"": """"," & vbLf &
        "        ""billing_street2"": """"," & vbLf & "        ""billing_phone"": """"," & vbLf &
        "        ""discount_code"": """"," & vbLf & "        ""bonus_campaign_id"": """"," & vbLf &
        "        ""gift"": """"," & vbLf & "        ""gift_message"": """"," & vbLf &
        "        ""merchant_id"": """"," & vbLf & "        ""merchant_created_at"": """"," & vbLf &
        "        ""merchant_country"": """"," & vbLf & "        ""merchant_category"": """"," & vbLf &
        "        ""details_url"": """"," & vbLf & "        ""custom_fields"": { " & vbLf &
        "            ""is_intangible_item"": true," & vbLf & "            ""is_pay_on_delivery"": true," & vbLf &
        "            ""departure_airport"": ""BUD""," & vbLf & "            ""days_to_board"": 1," & vbLf &
        "            ""arrival_airport"": ""MXP""" & vbLf & "	}" & vbLf & "}"





        Return returnValue
    End Function

    Public Shared Function GetSeonparametersFromSettings(stage As Integer) As String
        Dim returnValue As String

        Select Case stage
            Case 1 'url
                returnValue = ExecuteSQLScalarComandString("select apiLink from applicationsConfigurtion where id=6")
            Case 2 'apiKey
                returnValue = ExecuteSQLScalarComandString("select apiKey from applicationsConfigurtion where id=6")
        End Select

        Return returnValue
    End Function


    Public Shared Function SendSeonFraudAPIRequest(stage As Integer, userIDFor As String, ipForSend As String, emailFor As String, phone As String) As String

        Dim client = New RestClient(GetSeonparametersFromSettings(1))
        client.Timeout = -1
        Dim request = New RestRequest(Method.POST)
        request.AddHeader("X-API-KEY", GetSeonparametersFromSettings(2))
        request.AddHeader("Content-Type", "application/json")
        Dim body = CreateRequestBody(stage, userIDFor, ipForSend, emailFor, phone)
        request.AddParameter("application/json", body, ParameterType.RequestBody)
        Dim response As IRestResponse = client.Execute(request)


        Return response.Content
    End Function






End Class
