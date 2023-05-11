#product names and service 
#https://learn.microsoft.com/en-us/azure/active-directory/enterprise-users/licensing-service-plan-reference

#$MicrosoftDocumentationURI = 'https://learn.microsoft.com/en-us/azure/active-directory/enterprise-users/licensing-service-plan-reference'
#$MicrosoftDocumentationCSVURI = ((Invoke-WebRequest -UseBasicParsing -Uri $MicrosoftDocumentationURI).Links | Where-Object { $_.href -like 'http*' } | Where-Object { $_.href -like '*.csv' }).href
#$MicrosoftProducts = Invoke-RestMethod -Uri $MicrosoftDocumentationCSVURI -Method Get | ConvertFrom-CSV | Select-Object String_ID, Product_Display_Name -Unique

#$SearchSKU = 'ENTERPRISEPACK'
#$MicrosoftProducts | Where-Object { $_.String_ID -eq $SearchSKU } | Select-Object Product_Display_Name
