#Requires -Modules AzureAD, ExchangeOnlineManagement

#####Declarations#####
#Allow -Verbose to Work
[CmdletBinding()]
Param()
#Include GUI Elements in Script
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
[void] [System.Windows.Forms.Application]::EnableVisualStyles()
#Allow Quitbox to work
$quitboxOutput = ""
#Sku Lookup Table
$SkuToFriendly = @{
    "c42b9cae-ea4f-4ab7-9717-81576235ccac" = "DevPack E5 (No Teams or Audio)"
    "8f0c5670-4e56-4892-b06d-91c085d7004f" = "APP CONNECT IW"
    "0c266dff-15dd-4b49-8397-2bb16070ed52" = "Microsoft 365 Audio Conferencing"
    "2b9c8e7c-319c-43a2-a2a0-48c5c6161de7" = "AZURE ACTIVE DIRECTORY BASIC"
    "078d2b04-f1bd-4111-bbd4-b4b1b354cef4" = "AZURE ACTIVE DIRECTORY PREMIUM P1"
    "84a661c4-e949-4bd2-a560-ed7766fcaf2b" = "AZURE ACTIVE DIRECTORY PREMIUM P2"
    "c52ea49f-fe5d-4e95-93ba-1de91d380f89" = "AZURE INFORMATION PROTECTION PLAN 1"
    "295a8eb0-f78d-45c7-8b5b-1eed5ed02dff" = "COMMON AREA PHONE"
    "47794cd0-f0e5-45c5-9033-2eb6b5fc84e0" = "COMMUNICATIONS CREDITS"
    "ea126fc5-a19e-42e2-a731-da9d437bffcf" = "DYNAMICS 365 CUSTOMER ENGAGEMENT PLAN ENTERPRISE EDITION"
    "749742bf-0d37-4158-a120-33567104deeb" = "DYNAMICS 365 FOR CUSTOMER SERVICE ENTERPRISE EDITION"
    "cc13a803-544e-4464-b4e4-6d6169a138fa" = "DYNAMICS 365 FOR FINANCIALS BUSINESS EDITION"
    "8edc2cf8-6438-4fa9-b6e3-aa1660c640cc" = "DYNAMICS 365 FOR SALES AND CUSTOMER SERVICE ENTERPRISE EDITION"
    "1e1a282c-9c54-43a2-9310-98ef728faace" = "DYNAMICS 365 FOR SALES ENTERPRISE EDITION"
    "f2e48cb3-9da0-42cd-8464-4a54ce198ad0" = "DYNAMICS 365 FOR SUPPLY CHAIN MANAGEMENT"
    "8e7a3d30-d97d-43ab-837c-d7701cef83dc" = "DYNAMICS 365 FOR TEAM MEMBERS ENTERPRISE EDITION"
    "338148b6-1b11-4102-afb9-f92b6cdc0f8d" = "DYNAMICS 365 P1 TRIAL FOR INFORMATION WORKERS"
    "b56e7ccc-d5c7-421f-a23b-5c18bdbad7c0" = "DYNAMICS 365 TALENT: ONBOARD"
    "7ac9fe77-66b7-4e5e-9e46-10eed1cff547" = "DYNAMICS 365 TEAM MEMBERS"
    "ccba3cfe-71ef-423a-bd87-b6df3dce59a9" = "DYNAMICS 365 UNF OPS PLAN ENT EDITION"
    "efccb6f7-5641-4e0e-bd10-b4976e1bf68e" = "ENTERPRISE MOBILITY + SECURITY E3"
    "b05e124f-c7cc-45a0-a6aa-8cf78c946968" = "ENTERPRISE MOBILITY + SECURITY E5"
    "4b9405b0-7788-4568-add1-99614e613b69" = "EXCHANGE ONLINE (PLAN 1)"
    "19ec0d23-8335-4cbd-94ac-6050e30712fa" = "EXCHANGE ONLINE (PLAN 2)"
    "ee02fd1b-340e-4a4b-b355-4a514e4c8943" = "EXCHANGE ONLINE ARCHIVING FOR EXCHANGE ONLINE"
    "90b5e015-709a-4b8b-b08e-3200f994494c" = "EXCHANGE ONLINE ARCHIVING FOR EXCHANGE SERVER"
    "7fc0182e-d107-4556-8329-7caaa511197b" = "EXCHANGE ONLINE ESSENTIALS (ExO P1 BASED)"
    "e8f81a67-bd96-4074-b108-cf193eb9433b" = "EXCHANGE ONLINE ESSENTIALS"
    "80b2d799-d2ba-4d2a-8842-fb0d0f3a4b82" = "EXCHANGE ONLINE KIOSK"
    "cb0a98a8-11bc-494c-83d9-c1b1ac65327e" = "EXCHANGE ONLINE POP"
    "061f9ace-7d42-4136-88ac-31dc755f143f" = "INTUNE"
    "b17653a4-2443-4e8c-a550-18249dda78bb" = "Microsoft 365 A1"
    "4b590615-0888-425a-a965-b3bf7789848d" = "MICROSOFT 365 A3 FOR FACULTY"
    "7cfd9a2b-e110-4c39-bf20-c6a3f36a3121" = "MICROSOFT 365 A3 FOR STUDENTS"
    "e97c048c-37a4-45fb-ab50-922fbf07a370" = "MICROSOFT 365 A5 FOR FACULTY"
    "46c119d4-0379-4a9d-85e4-97c66d3f909e" = "MICROSOFT 365 A5 FOR STUDENTS"
    "cdd28e44-67e3-425e-be4c-737fab2899d3" = "MICROSOFT 365 APPS FOR BUSINESS"
    "b214fe43-f5a3-4703-beeb-fa97188220fc" = "MICROSOFT 365 APPS FOR BUSINESS"
    "c2273bd0-dff7-4215-9ef5-2c7bcfb06425" = "MICROSOFT 365 APPS FOR ENTERPRISE"
    "2d3091c7-0712-488b-b3d8-6b97bde6a1f5" = "MICROSOFT 365 AUDIO CONFERENCING FOR GCC"
    "3b555118-da6a-4418-894f-7df1e2096870" = "MICROSOFT 365 BUSINESS BASIC"
    "dab7782a-93b1-4074-8bb1-0e61318bea0b" = "MICROSOFT 365 BUSINESS BASIC"
    "f245ecc8-75af-4f8e-b61f-27d8114de5f3" = "MICROSOFT 365 BUSINESS STANDARD"
    "ac5cef5d-921b-4f97-9ef3-c99076e5470f" = "MICROSOFT 365 BUSINESS STANDARD - PREPAID LEGACY"
    "cbdc14ab-d96c-4c30-b9f4-6ada7cdc1d46" = "MICROSOFT 365 BUSINESS PREMIUM"
    "11dee6af-eca8-419f-8061-6864517c1875" = "MICROSOFT 365 DOMESTIC CALLING PLAN (120 Minutes)"
    "05e9a617-0261-4cee-bb44-138d3ef5d965" = "MICROSOFT 365 E3"
    "06ebc4ee-1bb5-47dd-8120-11324bc54e06" = "Microsoft 365 E5"
    "d61d61cc-f992-433f-a577-5bd016037eeb" = "Microsoft 365 E3_USGOV_DOD"
    "ca9d1dd9-dfe9-4fef-b97c-9bc1ea3c3658" = "Microsoft 365 E3_USGOV_GCCHIGH"
    "184efa21-98c3-4e5d-95ab-d07053a96e67" = "Microsoft 365 E5 Compliance"
    "26124093-3d78-432b-b5dc-48bf992543d5" = "Microsoft 365 E5 Security"
    "44ac31e7-2999-4304-ad94-c948886741d4" = "Microsoft 365 E5 Security for EMS E5"
    "44575883-256e-4a79-9da4-ebe9acabe2b2" = "Microsoft 365 F1"
    "66b55226-6b4f-492c-910c-a3b7a3c9d993" = "Microsoft 365 F3"
    "f30db892-07e9-47e9-837c-80727f46fd3d" = "MICROSOFT FLOW FREE"
    "e823ca47-49c4-46b3-b38d-ca11d5abe3d2" = "MICROSOFT 365 G3 GCC"
    "e43b5b99-8dfb-405f-9987-dc307f34bcbd" = "MICROSOFT 365 PHONE SYSTEM"
    "d01d9287-694b-44f3-bcc5-ada78c8d953e" = "MICROSOFT 365 PHONE SYSTEM FOR DOD"
    "d979703c-028d-4de5-acbf-7955566b69b9" = "MICROSOFT 365 PHONE SYSTEM FOR FACULTY"
    "a460366a-ade7-4791-b581-9fbff1bdaa85" = "MICROSOFT 365 PHONE SYSTEM FOR GCC"
    "7035277a-5e49-4abc-a24f-0ec49c501bb5" = "MICROSOFT 365 PHONE SYSTEM FOR GCCHIGH"
    "aa6791d3-bb09-4bc2-afed-c30c3fe26032" = "MICROSOFT 365 PHONE SYSTEM FOR SMALL AND MEDIUM BUSINESS"
    "1f338bbc-767e-4a1e-a2d4-b73207cc5b93" = "MICROSOFT 365 PHONE SYSTEM FOR STUDENTS"
    "ffaf2d68-1c95-4eb3-9ddd-59b81fba0f61" = "MICROSOFT 365 PHONE SYSTEM FOR TELSTRA"
    "b0e7de67-e503-4934-b729-53d595ba5cd1" = "MICROSOFT 365 PHONE SYSTEM_USGOV_DOD"
    "985fcb26-7b94-475b-b512-89356697be71" = "MICROSOFT 365 PHONE SYSTEM_USGOV_GCCHIGH"
    "440eaaa8-b3e0-484b-a8be-62870b9ba70a" = "MICROSOFT 365 PHONE SYSTEM - VIRTUAL USER"
    "2347355b-4e81-41a4-9c22-55057a399791" = "MICROSOFT 365 SECURITY AND COMPLIANCE FOR FLW"
    "726a0894-2c77-4d65-99da-9775ef05aad1" = "MICROSOFT BUSINESS CENTER"
    "111046dd-295b-4d6d-9724-d52ac90bd1f2" = "MICROSOFT DEFENDER FOR ENDPOINT"
    "906af65a-2970-46d5-9b58-4e9aa50f0657" = "MICROSOFT DYNAMICS CRM ONLINE BASIC"
    "d17b27af-3f49-4822-99f9-56a661538792" = "MICROSOFT DYNAMICS CRM ONLINE"
    "ba9a34de-4489-469d-879c-0f0f145321cd" = "MS IMAGINE ACADEMY"
    "2c21e77a-e0d6-4570-b38a-7ff2dc17d2ca" = "MICROSOFT INTUNE DEVICE FOR GOVERNMENT"
    "dcb1a3ae-b33f-4487-846a-a640262fadf4" = "MICROSOFT POWER APPS PLAN 2 TRIAL"
    "e6025b08-2fa5-4313-bd0a-7e5ffca32958" = "MICROSOFT INTUNE SMB"
    "1f2f344a-700d-42c9-9427-5cea1d5d7ba6" = "MICROSOFT STREAM"
    "16ddbbfc-09ea-4de2-b1d7-312db6112d70" = "MICROSOFT TEAM (FREE)"
    "710779e8-3d4a-4c88-adb9-386c958d1fdf" = "MICROSOFT TEAMS EXPLORATORY"
    "a4585165-0533-458a-97e3-c400570268c4" = "Office 365 A5 for faculty"
    "ee656612-49fa-43e5-b67e-cb1fdf7699df" = "Office 365 A5 for students"
    "1b1b1f7a-8355-43b6-829f-336cfccb744c" = "Office 365 Advanced Compliance"
    "4ef96642-f096-40de-a3e9-d83fb2f90211" = "Microsoft Defender for Office 365 (Plan 1)"
    "18181a46-0d4e-45cd-891e-60aabd171b4e" = "OFFICE 365 E1"
    "6634e0ce-1a9f-428c-a498-f84ec7b8aa2e" = "OFFICE 365 E2"
    "6fd2c87f-b296-42f0-b197-1e91e994b900" = "OFFICE 365 E3"
    "189a915c-fe4f-4ffa-bde4-85b9628d07a0" = "OFFICE 365 E3 DEVELOPER"
    "b107e5a3-3e60-4c0d-a184-a7e4395eb44c" = "Office 365 E3_USGOV_DOD"
    "aea38a85-9bd5-4981-aa00-616b411205bf" = "Office 365 E3_USGOV_GCCHIGH"
    "1392051d-0cb9-4b7a-88d5-621fee5e8711" = "OFFICE 365 E4"
    "c7df2760-2c81-4ef7-b578-5b5392b571df" = "OFFICE 365 E5"
    "26d45bd9-adf1-46cd-a9e1-51e9a5524128" = "OFFICE 365 E5 WITHOUT AUDIO CONFERENCING"
    "4b585984-651b-448a-9e53-3b10f069cf7f" = "OFFICE 365 F3"
    "535a3a29-c5f0-42fe-8215-d3b9e1f38c4a" = "OFFICE 365 G3 GCC"
    "04a7fb0d-32e0-4241-b4f5-3f7618cd1162" = "OFFICE 365 MIDSIZE BUSINESS"
    "bd09678e-b83c-4d3f-aaba-3dad4abd128b" = "OFFICE 365 SMALL BUSINESS"
    "fc14ec4a-4169-49a4-a51e-2c852931814b" = "OFFICE 365 SMALL BUSINESS PREMIUM"
    "e6778190-713e-4e4f-9119-8b8238de25df" = "ONEDRIVE FOR BUSINESS (PLAN 1)"
    "ed01faf2-1d88-4947-ae91-45ca18703a96" = "ONEDRIVE FOR BUSINESS (PLAN 2)"
    "87bbbc60-4754-4998-8c88-227dca264858" = "POWERAPPS AND LOGIC FLOWS"
    "a403ebcc-fae0-4ca2-8c8c-7a907fd6c235" = "POWER BI (FREE)"
    "45bc2c81-6072-436a-9b0b-3b12eefbc402" = "POWER BI FOR OFFICE 365 ADD-ON"
    "f8a1db68-be16-40ed-86d5-cb42ce701560" = "POWER BI PRO"
    "a10d5e58-74da-4312-95c8-76be4e5b75a0" = "PROJECT FOR OFFICE 365"
    "776df282-9fc0-4862-99e2-70e561b9909e" = "PROJECT ONLINE ESSENTIALS"
    "09015f9f-377f-4538-bbb5-f75ceb09358a" = "PROJECT ONLINE PREMIUM"
    "2db84718-652c-47a7-860c-f10d8abbdae3" = "PROJECT ONLINE PREMIUM WITHOUT PROJECT CLIENT"
    "53818b1b-4a27-454b-8896-0dba576410e6" = "PROJECT ONLINE PROFESSIONAL"
    "f82a60b8-1ee3-4cfb-a4fe-1c6a53c2656c" = "PROJECT ONLINE WITH PROJECT FOR OFFICE 365"
    "beb6439c-caad-48d3-bf46-0c82871e12be" = "PROJECT PLAN 1"
    "1fc08a02-8b3d-43b9-831e-f76859e04e1a" = "SHAREPOINT ONLINE (PLAN 1)"
    "a9732ec9-17d9-494c-a51c-d6b45b384dcb" = "SHAREPOINT ONLINE (PLAN 2)"
    "b8b749f8-a4ef-4887-9539-c95b1eaa5db7" = "SKYPE FOR BUSINESS ONLINE (PLAN 1)"
    "d42c793f-6c78-4f43-92ca-e8f6a02b035f" = "SKYPE FOR BUSINESS ONLINE (PLAN 2)"
    "d3b4fe1f-9992-4930-8acb-ca6ec609365e" = "SKYPE FOR BUSINESS PSTN DOMESTIC AND INTERNATIONAL CALLING"
    "0dab259f-bf13-4952-b7f8-7db8f131b28d" = "SKYPE FOR BUSINESS PSTN DOMESTIC CALLING"
    "54a152dc-90de-4996-93d2-bc47e670fc06" = "SKYPE FOR BUSINESS PSTN DOMESTIC CALLING (120 Minutes)"
    "4016f256-b063-4864-816e-d818aad600c9" = "TOPIC EXPERIENCES"
    "de3312e1-c7b0-46e6-a7c3-a515ff90bc86" = "TELSTRA CALLING FOR O365"
    "4b244418-9658-4451-a2b8-b5e2b364e9bd" = "VISIO ONLINE PLAN 1"
    "c5928f49-12ba-48f7-ada3-0d743a3601d5" = "VISIO ONLINE PLAN 2"
    "4ae99959-6b0f-43b0-b1ce-68146001bdba" = "VISIO PLAN 2 FOR GCC"
    "cb10e6cd-9da4-4992-867b-67546b1db821" = "WINDOWS 10 ENTERPRISE E3"
    "6a0f6da5-0b87-4190-a6ae-9bb5a2b9546a" = "WINDOWS 10 ENTERPRISE E3"
    "488ba24a-39a9-4473-8ee5-19291e71b002" = "Windows 10 Enterprise E5"
    "6470687e-a428-4b7a-bef2-8a291ad947c9" = "WINDOWS STORE FOR BUSINESS"
}
#Sku Reverse Lookup Table
$FriendlyToSku = @{
    "DevPack E5 (No Teams or Audio)" = "c42b9cae-ea4f-4ab7-9717-81576235ccac"
    "APP CONNECT IW" = "8f0c5670-4e56-4892-b06d-91c085d7004f"
    "Microsoft 365 Audio Conferencing" = "0c266dff-15dd-4b49-8397-2bb16070ed52"
    "AZURE ACTIVE DIRECTORY BASIC" = "2b9c8e7c-319c-43a2-a2a0-48c5c6161de7"
    "AZURE ACTIVE DIRECTORY PREMIUM P1" = "078d2b04-f1bd-4111-bbd4-b4b1b354cef4"
    "AZURE ACTIVE DIRECTORY PREMIUM P2" = "84a661c4-e949-4bd2-a560-ed7766fcaf2b"
    "AZURE INFORMATION PROTECTION PLAN 1" = "c52ea49f-fe5d-4e95-93ba-1de91d380f89"
    "COMMON AREA PHONE" = "295a8eb0-f78d-45c7-8b5b-1eed5ed02dff"
    "COMMUNICATIONS CREDITS" = "47794cd0-f0e5-45c5-9033-2eb6b5fc84e0"
    "DYNAMICS 365 CUSTOMER ENGAGEMENT PLAN ENTERPRISE EDITION" = "ea126fc5-a19e-42e2-a731-da9d437bffcf"
    "DYNAMICS 365 FOR CUSTOMER SERVICE ENTERPRISE EDITION" = "749742bf-0d37-4158-a120-33567104deeb"
    "DYNAMICS 365 FOR FINANCIALS BUSINESS EDITION" = "cc13a803-544e-4464-b4e4-6d6169a138fa"
    "DYNAMICS 365 FOR SALES AND CUSTOMER SERVICE ENTERPRISE EDITION" = "8edc2cf8-6438-4fa9-b6e3-aa1660c640cc"
    "DYNAMICS 365 FOR SALES ENTERPRISE EDITION" = "1e1a282c-9c54-43a2-9310-98ef728faace"
    "DYNAMICS 365 FOR SUPPLY CHAIN MANAGEMENT" = "f2e48cb3-9da0-42cd-8464-4a54ce198ad0"
    "DYNAMICS 365 FOR TEAM MEMBERS ENTERPRISE EDITION" = "8e7a3d30-d97d-43ab-837c-d7701cef83dc"
    "DYNAMICS 365 P1 TRIAL FOR INFORMATION WORKERS" = "338148b6-1b11-4102-afb9-f92b6cdc0f8d"
    "DYNAMICS 365 TALENT: ONBOARD" = "b56e7ccc-d5c7-421f-a23b-5c18bdbad7c0"
    "DYNAMICS 365 TEAM MEMBERS" = "7ac9fe77-66b7-4e5e-9e46-10eed1cff547"
    "DYNAMICS 365 UNF OPS PLAN ENT EDITION" = "ccba3cfe-71ef-423a-bd87-b6df3dce59a9"
    "ENTERPRISE MOBILITY + SECURITY E3" = "efccb6f7-5641-4e0e-bd10-b4976e1bf68e"
    "ENTERPRISE MOBILITY + SECURITY E5" = "b05e124f-c7cc-45a0-a6aa-8cf78c946968"
    "EXCHANGE ONLINE (PLAN 1)" = "4b9405b0-7788-4568-add1-99614e613b69"
    "EXCHANGE ONLINE (PLAN 2)" = "19ec0d23-8335-4cbd-94ac-6050e30712fa"
    "EXCHANGE ONLINE ARCHIVING FOR EXCHANGE ONLINE" = "ee02fd1b-340e-4a4b-b355-4a514e4c8943"
    "EXCHANGE ONLINE ARCHIVING FOR EXCHANGE SERVER" = "90b5e015-709a-4b8b-b08e-3200f994494c"
    "EXCHANGE ONLINE ESSENTIALS (ExO P1 BASED)" = "7fc0182e-d107-4556-8329-7caaa511197b"
    "EXCHANGE ONLINE ESSENTIALS" = "e8f81a67-bd96-4074-b108-cf193eb9433b"
    "EXCHANGE ONLINE KIOSK" = "80b2d799-d2ba-4d2a-8842-fb0d0f3a4b82"
    "EXCHANGE ONLINE POP" = "cb0a98a8-11bc-494c-83d9-c1b1ac65327e"
    "INTUNE" = "061f9ace-7d42-4136-88ac-31dc755f143f"
    "Microsoft 365 A1" = "b17653a4-2443-4e8c-a550-18249dda78bb"
    "MICROSOFT 365 A3 FOR FACULTY" = "4b590615-0888-425a-a965-b3bf7789848d"
    "MICROSOFT 365 A3 FOR STUDENTS" = "7cfd9a2b-e110-4c39-bf20-c6a3f36a3121"
    "MICROSOFT 365 A5 FOR FACULTY" = "e97c048c-37a4-45fb-ab50-922fbf07a370"
    "MICROSOFT 365 A5 FOR STUDENTS" = "46c119d4-0379-4a9d-85e4-97c66d3f909e"
    "MICROSOFT 365 APPS FOR BUSINESS" = "cdd28e44-67e3-425e-be4c-737fab2899d3"
    "MICROSOFT 365 APPS FOR BUSINESS (2)" = "b214fe43-f5a3-4703-beeb-fa97188220fc"
    "MICROSOFT 365 APPS FOR ENTERPRISE" = "c2273bd0-dff7-4215-9ef5-2c7bcfb06425"
    "MICROSOFT 365 AUDIO CONFERENCING FOR GCC" = "2d3091c7-0712-488b-b3d8-6b97bde6a1f5"
    "MICROSOFT 365 BUSINESS BASIC" = "3b555118-da6a-4418-894f-7df1e2096870"
    "MICROSOFT 365 BUSINESS BASIC (2)" = "dab7782a-93b1-4074-8bb1-0e61318bea0b"
    "MICROSOFT 365 BUSINESS STANDARD" = "f245ecc8-75af-4f8e-b61f-27d8114de5f3"
    "MICROSOFT 365 BUSINESS STANDARD - PREPAID LEGACY" = "ac5cef5d-921b-4f97-9ef3-c99076e5470f"
    "MICROSOFT 365 BUSINESS PREMIUM" = "cbdc14ab-d96c-4c30-b9f4-6ada7cdc1d46"
    "MICROSOFT 365 DOMESTIC CALLING PLAN (120 Minutes)" = "11dee6af-eca8-419f-8061-6864517c1875"
    "MICROSOFT 365 E3" = "05e9a617-0261-4cee-bb44-138d3ef5d965"
    "Microsoft 365 E5" = "06ebc4ee-1bb5-47dd-8120-11324bc54e06"
    "Microsoft 365 E3_USGOV_DOD" = "d61d61cc-f992-433f-a577-5bd016037eeb"
    "Microsoft 365 E3_USGOV_GCCHIGH" = "ca9d1dd9-dfe9-4fef-b97c-9bc1ea3c3658"
    "Microsoft 365 E5 Compliance" = "184efa21-98c3-4e5d-95ab-d07053a96e67"
    "Microsoft 365 E5 Security" = "26124093-3d78-432b-b5dc-48bf992543d5"
    "Microsoft 365 E5 Security for EMS E5" = "44ac31e7-2999-4304-ad94-c948886741d4"
    "Microsoft 365 F1" = "44575883-256e-4a79-9da4-ebe9acabe2b2"
    "Microsoft 365 F3" = "66b55226-6b4f-492c-910c-a3b7a3c9d993"
    "MICROSOFT FLOW FREE" = "f30db892-07e9-47e9-837c-80727f46fd3d"
    "MICROSOFT 365 G3 GCC" = "e823ca47-49c4-46b3-b38d-ca11d5abe3d2"
    "MICROSOFT 365 PHONE SYSTEM" = "e43b5b99-8dfb-405f-9987-dc307f34bcbd"
    "MICROSOFT 365 PHONE SYSTEM FOR DOD" = "d01d9287-694b-44f3-bcc5-ada78c8d953e"
    "MICROSOFT 365 PHONE SYSTEM FOR FACULTY" = "d979703c-028d-4de5-acbf-7955566b69b9"
    "MICROSOFT 365 PHONE SYSTEM FOR GCC" = "a460366a-ade7-4791-b581-9fbff1bdaa85"
    "MICROSOFT 365 PHONE SYSTEM FOR GCCHIGH" = "7035277a-5e49-4abc-a24f-0ec49c501bb5"
    "MICROSOFT 365 PHONE SYSTEM FOR SMALL AND MEDIUM BUSINESS" = "aa6791d3-bb09-4bc2-afed-c30c3fe26032"
    "MICROSOFT 365 PHONE SYSTEM FOR STUDENTS" = "1f338bbc-767e-4a1e-a2d4-b73207cc5b93"
    "MICROSOFT 365 PHONE SYSTEM FOR TELSTRA" = "ffaf2d68-1c95-4eb3-9ddd-59b81fba0f61"
    "MICROSOFT 365 PHONE SYSTEM_USGOV_DOD" = "b0e7de67-e503-4934-b729-53d595ba5cd1"
    "MICROSOFT 365 PHONE SYSTEM_USGOV_GCCHIGH" = "985fcb26-7b94-475b-b512-89356697be71"
    "MICROSOFT 365 PHONE SYSTEM - VIRTUAL USER" = "440eaaa8-b3e0-484b-a8be-62870b9ba70a"
    "MICROSOFT 365 SECURITY AND COMPLIANCE FOR FLW" = "2347355b-4e81-41a4-9c22-55057a399791"
    "MICROSOFT BUSINESS CENTER" = "726a0894-2c77-4d65-99da-9775ef05aad1"
    "MICROSOFT DEFENDER FOR ENDPOINT" = "111046dd-295b-4d6d-9724-d52ac90bd1f2"
    "MICROSOFT DYNAMICS CRM ONLINE BASIC" = "906af65a-2970-46d5-9b58-4e9aa50f0657"
    "MICROSOFT DYNAMICS CRM ONLINE" = "d17b27af-3f49-4822-99f9-56a661538792"
    "MS IMAGINE ACADEMY" = "ba9a34de-4489-469d-879c-0f0f145321cd"
    "MICROSOFT INTUNE DEVICE FOR GOVERNMENT" = "2c21e77a-e0d6-4570-b38a-7ff2dc17d2ca"
    "MICROSOFT POWER APPS PLAN 2 TRIAL" = "dcb1a3ae-b33f-4487-846a-a640262fadf4"
    "MICROSOFT INTUNE SMB" = "e6025b08-2fa5-4313-bd0a-7e5ffca32958"
    "MICROSOFT STREAM" = "1f2f344a-700d-42c9-9427-5cea1d5d7ba6"
    "MICROSOFT TEAM (FREE)" = "16ddbbfc-09ea-4de2-b1d7-312db6112d70"
    "MICROSOFT TEAMS EXPLORATORY" = "710779e8-3d4a-4c88-adb9-386c958d1fdf"
    "Office 365 A5 for faculty" = "a4585165-0533-458a-97e3-c400570268c4"
    "Office 365 A5 for students" = "ee656612-49fa-43e5-b67e-cb1fdf7699df"
    "Office 365 Advanced Compliance" = "1b1b1f7a-8355-43b6-829f-336cfccb744c"
    "Microsoft Defender for Office 365 (Plan 1)" = "4ef96642-f096-40de-a3e9-d83fb2f90211"
    "OFFICE 365 E1" = "18181a46-0d4e-45cd-891e-60aabd171b4e"
    "OFFICE 365 E2" = "6634e0ce-1a9f-428c-a498-f84ec7b8aa2e"
    "OFFICE 365 E3" = "6fd2c87f-b296-42f0-b197-1e91e994b900"
    "OFFICE 365 E3 DEVELOPER" = "189a915c-fe4f-4ffa-bde4-85b9628d07a0"
    "Office 365 E3_USGOV_DOD" = "b107e5a3-3e60-4c0d-a184-a7e4395eb44c"
    "Office 365 E3_USGOV_GCCHIGH" = "aea38a85-9bd5-4981-aa00-616b411205bf"
    "OFFICE 365 E4" = "1392051d-0cb9-4b7a-88d5-621fee5e8711"
    "OFFICE 365 E5" = "c7df2760-2c81-4ef7-b578-5b5392b571df"
    "OFFICE 365 E5 WITHOUT AUDIO CONFERENCING" = "26d45bd9-adf1-46cd-a9e1-51e9a5524128"
    "OFFICE 365 F3" = "4b585984-651b-448a-9e53-3b10f069cf7f"
    "OFFICE 365 G3 GCC" = "535a3a29-c5f0-42fe-8215-d3b9e1f38c4a"
    "OFFICE 365 MIDSIZE BUSINESS" = "04a7fb0d-32e0-4241-b4f5-3f7618cd1162"
    "OFFICE 365 SMALL BUSINESS" = "bd09678e-b83c-4d3f-aaba-3dad4abd128b"
    "OFFICE 365 SMALL BUSINESS PREMIUM" = "fc14ec4a-4169-49a4-a51e-2c852931814b"
    "ONEDRIVE FOR BUSINESS (PLAN 1)" = "e6778190-713e-4e4f-9119-8b8238de25df"
    "ONEDRIVE FOR BUSINESS (PLAN 2)" = "ed01faf2-1d88-4947-ae91-45ca18703a96"
    "POWERAPPS AND LOGIC FLOWS" = "87bbbc60-4754-4998-8c88-227dca264858"
    "POWER BI (FREE)" = "a403ebcc-fae0-4ca2-8c8c-7a907fd6c235"
    "POWER BI FOR OFFICE 365 ADD-ON" = "45bc2c81-6072-436a-9b0b-3b12eefbc402"
    "POWER BI PRO" = "f8a1db68-be16-40ed-86d5-cb42ce701560"
    "PROJECT FOR OFFICE 365" = "a10d5e58-74da-4312-95c8-76be4e5b75a0"
    "PROJECT ONLINE ESSENTIALS" = "776df282-9fc0-4862-99e2-70e561b9909e"
    "PROJECT ONLINE PREMIUM" = "09015f9f-377f-4538-bbb5-f75ceb09358a"
    "PROJECT ONLINE PREMIUM WITHOUT PROJECT CLIENT" = "2db84718-652c-47a7-860c-f10d8abbdae3"
    "PROJECT ONLINE PROFESSIONAL" = "53818b1b-4a27-454b-8896-0dba576410e6"
    "PROJECT ONLINE WITH PROJECT FOR OFFICE 365" = "f82a60b8-1ee3-4cfb-a4fe-1c6a53c2656c"
    "PROJECT PLAN 1" = "beb6439c-caad-48d3-bf46-0c82871e12be"
    "SHAREPOINT ONLINE (PLAN 1)" = "1fc08a02-8b3d-43b9-831e-f76859e04e1a"
    "SHAREPOINT ONLINE (PLAN 2)" = "a9732ec9-17d9-494c-a51c-d6b45b384dcb"
    "SKYPE FOR BUSINESS ONLINE (PLAN 1)" = "b8b749f8-a4ef-4887-9539-c95b1eaa5db7"
    "SKYPE FOR BUSINESS ONLINE (PLAN 2)" = "d42c793f-6c78-4f43-92ca-e8f6a02b035f"
    "SKYPE FOR BUSINESS PSTN DOMESTIC AND INTERNATIONAL CALLING" = "d3b4fe1f-9992-4930-8acb-ca6ec609365e"
    "SKYPE FOR BUSINESS PSTN DOMESTIC CALLING" = "0dab259f-bf13-4952-b7f8-7db8f131b28d"
    "SKYPE FOR BUSINESS PSTN DOMESTIC CALLING (120 Minutes)" = "54a152dc-90de-4996-93d2-bc47e670fc06"
    "TOPIC EXPERIENCES" = "4016f256-b063-4864-816e-d818aad600c9"
    "TELSTRA CALLING FOR O365" = "de3312e1-c7b0-46e6-a7c3-a515ff90bc86"
    "VISIO ONLINE PLAN 1" = "4b244418-9658-4451-a2b8-b5e2b364e9bd"
    "VISIO ONLINE PLAN 2" = "c5928f49-12ba-48f7-ada3-0d743a3601d5"
    "VISIO PLAN 2 FOR GCC" = "4ae99959-6b0f-43b0-b1ce-68146001bdba"
    "WINDOWS 10 ENTERPRISE E3" = "cb10e6cd-9da4-4992-867b-67546b1db821"
    "WINDOWS 10 ENTERPRISE E3 (2)" = "6a0f6da5-0b87-4190-a6ae-9bb5a2b9546a"
    "Windows 10 Enterprise E5" = "488ba24a-39a9-4473-8ee5-19291e71b002"
    "WINDOWS STORE FOR BUSINESS" = "6470687e-a428-4b7a-bef2-8a291ad947c9"
}
#Usage Location Lookup Table
$UsageLocations=@{
    "United States" = "US"
    "United Kingdom" = "UK"
}

#Function to Check If Mailbox Exists Before Touching It
function MailboxExistCheck {
    Write-Verbose "Checking If Mailbox Exists"
    Clear-Variable MailboxExistsCheck -ErrorAction SilentlyContinue
    #Start Mailbox Check Wait Loop
    while ($MailboxExistsCheck -ne "YES") {
        try {
            Get-EXOMailbox $UPN -ErrorAction Throw | Out-Null
            $MailboxExistsCheck = "YES"
        }
        catch {
            Write-Verbose "Mailbox Does Not Exist, Waiting 60 Seconds and Trying Again"
            Start-Sleep -Seconds 60
            $MailboxExistsCheck = "NO"
        }
    }#End Mailbox Check Wait Loop    
}
#Verification Check to enable OK button
function CheckAllBoxes{
    if ( $passwordTextbox.Text.Length -and ($domainComboBox.SelectedIndex -ge 0) -and $usernameTextbox.Text.Length -and $firstnameTextbox.Text.Length -and $lastnameTextbox.Text.Length )
    {
        $okButton.Enabled = $true
    }
    else {
        $okButton.Enabled = $false
    }
}
#####End of Declarations#####

# Test And Connect To AzureAD If Needed
try {
    Write-Verbose -Message "Testing connection to Azure AD"
    Get-AzureAdDomain -ErrorAction Stop | Out-Null
    Write-Verbose -Message "Already connected to Azure AD"
}
catch {
    Write-Verbose -Message "Connecting to Azure AD"
    Connect-AzureAD
}
#Test And Connect To Microsoft Exchange Online If Needed
try {
    Write-Verbose -Message "Testing connection to Microsoft Exchange Online"
    Get-Mailbox -ErrorAction Stop | Out-Null
    Write-Verbose -Message "Already connected to Microsoft Exchange Online"
}
catch {
    Write-Verbose -Message "Connecting to Microsoft Exchange Online"
    Connect-ExchangeOnline
}

#Start While Loop for Quitbox
while ($quitboxOutput -ne "NO"){
    #####Create User Details Form#####
    Clear-Variable passwordTextbox -ErrorAction SilentlyContinue
    Clear-Variable domainComboBox -ErrorAction SilentlyContinue
    Clear-Variable usernameTextbox -ErrorAction SilentlyContinue
    Clear-Variable lastnameTextbox -ErrorAction SilentlyContinue
    Clear-Variable firstnameTextbox -ErrorAction SilentlyContinue
    Clear-Variable displayname -ErrorAction SilentlyContinue
    Clear-Variable PasswordProfile -ErrorAction SilentlyContinue
    Clear-Variable UPN -ErrorAction SilentlyContinue
        
    $userdetailForm = New-Object System.Windows.Forms.Form
    $userdetailForm.Text = "User Details"
    $userdetailForm.Autosize = $true
    $userdetailForm.StartPosition = 'CenterScreen'
    $userdetailForm.Topmost = $true

    $okButton = New-Object System.Windows.Forms.Button
    $okButton.TabIndex = 6
    $okbutton.Dock = [System.Windows.Forms.DockStyle]::Bottom
    $okButton.Text = 'OK'
    $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $okButton.Enabled = $false
    $userdetailForm.AcceptButton = $okButton
    
    $userdetailForm.Controls.Add($okButton)

    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelbutton.TabIndex = 7
    $cancelButton.Dock = [System.Windows.Forms.DockStyle]::Bottom
    $cancelButton.Text = 'Cancel'
    $cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $userdetailForm.CancelButton = $cancelButton
    $userdetailForm.Controls.Add($cancelButton)

    $UsageLocationComboBox = New-Object System.Windows.Forms.ComboBox
    $UsageLocationCombobox.TabIndex = 5
    $UsageLocationComboBox.Dock = [System.Windows.Forms.DockStyle]::Top
    $UsageLocationComboBox.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    $UsageLocationComboBox.FormattingEnabled = $true
    foreach($UsageLocation in $UsageLocations.keys)
    {
        $null = $UsageLocationComboBox.Items.Add($usagelocation)
    }
    $UsageLocationComboBox.SelectedIndex = 0
    $UsageLocationComboBox.Add_TextChanged({CheckAllBoxes})
    $userdetailForm.Controls.Add($UsageLocationComboBox)
    
    $passwordTextbox = New-Object System.Windows.Forms.TextBox
    $passwordTextbox.TabIndex = 4
    $passwordTextbox.Dock = [System.Windows.Forms.DockStyle]::Top
    $passwordTextbox.Add_TextChanged({CheckAllBoxes})
    $userdetailForm.Controls.Add($passwordTextbox)
    
    $passwordLabel = New-Object System.Windows.Forms.Label
    $passwordLabel.Dock = [System.Windows.Forms.DockStyle]::Top
    $passwordLabel.Text = "Password"
    $userdetailForm.Controls.Add($passwordLabel)

    $domainComboBox = New-Object System.Windows.Forms.ComboBox
    $domainComboBox.TabIndex = 3
    $domainComboBox.Dock = [System.Windows.Forms.DockStyle]::Top
    $domainComboBox.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    $domainComboBox.FormattingEnabled = $true
    foreach($domain in Get-AzureADDomain){
        $null = $domainComboBox.Items.add($domain.Name)
    }
    $domainCombobox.Add_TextChanged({CheckAllBoxes})
    $userdetailForm.Controls.Add($domainComboBox)
    
    $domainLabel = New-Object System.Windows.Forms.Label
    $domainLabel.Dock = [System.Windows.Forms.DockStyle]::Top
    $domainLabel.Text = "@"
    $userdetailForm.Controls.Add($domainLabel)

    $usernameTextbox = New-Object System.Windows.Forms.TextBox
    $usernameTextbox.TabIndex = 2
    $usernameTextbox.Dock = [System.Windows.Forms.DockStyle]::Top
    $usernameTextbox.Add_TextChanged({CheckAllBoxes})
    $userdetailForm.Controls.Add($usernameTextbox)

    $usernameLabel = New-Object System.Windows.Forms.Label
    $usernameLabel.Dock = [System.Windows.Forms.DockStyle]::Top
    $usernameLabel.Text = "Username"
    $userdetailForm.Controls.Add($usernameLabel)

    $lastnameTextbox = New-Object System.Windows.Forms.TextBox
    $lastnameTextbox.TabIndex = 1
    $lastnameTextbox.Dock = [System.Windows.Forms.DockStyle]::Top
    $lastnameTextbox.Add_TextChanged({CheckAllBoxes})
    $userdetailForm.Controls.Add($lastnameTextbox)
    
    $lastnameLabel = New-Object System.Windows.Forms.Label
    $lastnameLabel.Dock = [System.Windows.Forms.DockStyle]::Top
    $lastnameLabel.Text = "Last Name"
    $userdetailForm.Controls.Add($lastnameLabel)

    $firstnameTextbox = New-Object System.Windows.Forms.TextBox
    $firstnameTextbox.TabIndex = 0
    $firstnameTextbox.Dock = [System.Windows.Forms.DockStyle]::Top
    $firstnameTextbox.Add_TextChanged({CheckAllBoxes})
    $userdetailForm.Controls.Add($firstnameTextbox)

    $firstnameLabel = New-Object System.Windows.Forms.Label
    $firstnameLabel.Dock = [System.Windows.Forms.DockStyle]::Top
    $firstnameLabel.Text = "First Name"
    $userdetailForm.Controls.Add($firstnameLabel)

    $userdetailForm.Add_Shown({$firstnameTextbox.Select()})
    $result = $userdetailForm.ShowDialog()
    #####End User Details Form#####

    #####User Detail Functionality#####
    #Create variables if OK button is pressed
    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        $PasswordProfile = New-Object -TypeName Microsoft.Open.AzureAD.Model.PasswordProfile
        $PasswordProfile.Password = $passwordTextbox.text
        $UPN = $usernameTextbox.Text + "@" + $domainCombobox.Text
        $displayname = $firstnameTextbox.text + " " + $lastnameTextbox.Text
        $usageloc = $UsageLocations[$UsageLocationComboBox.Text]
    }
    else { 
        Write-Verbose "Exiting Script." 
        Break 
    }
    #Test if User with matching UPN already exists, then exits
    try {
        Write-Verbose -Message "Testing If User Exists"
        Get-AzureAdUSer -ObjectID $UPN -ErrorAction Stop | Out-Null
        Write-Verbose -Message "User Exists, Exiting"
        Break
    }
    #Otherwise, creates user
    catch {
        Write-Verbose -Message "Creating User"
        $null = New-AzureADUser -DisplayName $displayname -PasswordProfile $PasswordProfile -UserPrincipalName $UPN -AccountEnabled $true -MailNickName "$($usernameTextbox.Text)" -USageLocation $usageloc
    }
    #####End User Detail Functionality#####
    
    #####Start Extra User Detail Form#####
    $additionaldetailForm = New-Object System.Windows.Forms.Form
    $additionaldetailForm.Text = "Addtional Details"
    $additionaldetailForm.Autosize = $true
    $additionaldetailForm.StartPosition = 'CenterScreen'
    $additionaldetailForm.Topmost = $true

    $okButton = New-Object System.Windows.Forms.Button
    $okButton.TabIndex = 3
    $okbutton.Dock = [System.Windows.Forms.DockStyle]::Bottom
    $okButton.Text = 'OK'
    $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $additionaldetailForm.AcceptButton = $okButton
    $additionaldetailForm.Controls.Add($okButton)

    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelbutton.TabIndex = 4
    $cancelButton.Dock = [System.Windows.Forms.DockStyle]::Bottom
    $cancelButton.Text = 'Cancel'
    $cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $additionaldetailForm.CancelButton = $cancelButton
    $additionaldetailForm.Controls.Add($cancelButton)

    $CustomAttribute1Textbox = New-Object System.Windows.Forms.TextBox
    $CustomAttribute1Textbox.TabIndex = 2
    $CustomAttribute1Textbox.Dock = [System.Windows.Forms.DockStyle]::Top
    $additionaldetailForm.Controls.Add($CustomAttribute1Textbox)

    $CustomAttribute1Label = New-Object System.Windows.Forms.Label
    $CustomAttribute1Label.Dock = [System.Windows.Forms.DockStyle]::Top
    $CustomAttribute1Label.Text = "Custom Attribute 1"
    $additionaldetailForm.Controls.Add($CustomAttribute1Label)

    $stateTextbox = New-Object System.Windows.Forms.TextBox
    $stateTextbox.TabIndex = 1
    $stateTextbox.Dock = [System.Windows.Forms.DockStyle]::Top
    $additionaldetailForm.Controls.Add($stateTextbox)

    $stateLabel = New-Object System.Windows.Forms.Label
    $stateLabel.Dock = [System.Windows.Forms.DockStyle]::Top
    $stateLabel.Text = "State"
    $additionaldetailForm.Controls.Add($stateLabel)

    $cityTextbox = New-Object System.Windows.Forms.TextBox
    $cityTextbox.TabIndex = 0
    $cityTextbox.Dock = [System.Windows.Forms.DockStyle]::Top
    $additionaldetailForm.Controls.Add($cityTextbox)

    $cityLabel = New-Object System.Windows.Forms.Label
    $cityLabel.Dock = [System.Windows.Forms.DockStyle]::Top
    $cityLabel.Text = "City"
    $additionaldetailForm.Controls.Add($cityLabel)

    $result = $additionaldetailForm.ShowDialog()
    #####End Extra Details Form#####

    #####Extra Detail Functionality#####
    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        if($cityTextbox.Text){
            Set-AzureADUser -ObjectId $UPN -City $cityTextbox.Text
        }
        if($stateTextbox.Text){
            Set-AzureADUser -ObjectID $UPN -State $stateTextbox.Text
        }
        if($CustomAttribute1Textbox.Text){
            MailboxExistCheck            
            Set-Mailbox $UPN -CustomAttribute1 $CustomAttribute1Textbox.Text
            Write-Verbose "Mailbox Exists, Added CustomAttribute1"
        }
    }
    else { 
        Write-Verbose "Exiting Script." 
        Break
    }
    #####End Extra Details Functionality#####    
    
    ##### Create License Selection Form #####
    $LicenseSelectWindow = New-Object System.Windows.Forms.Form
    $LicenseSelectWindow.Text = "Select Licenses"
    $LicenseSelectWindow.AutoSize = $true
    $LicenseSelectWindow.AutoSizeMode = "GrowAndShrink"
    $LicenseSelectWindow.MinimizeBox = $false
    $LicenseSelectWindow.MaximizeBox = $false
    $LicenseSelectWindow.StartPosition = "CenterScreen"
    $LicenseSelectWindow.FormBorderStyle = "Fixed3D"

    $CheckedListBox = New-Object System.Windows.Forms.CheckedListBox
    $CheckedListBox.AutoSize = $true
    $CheckedListBox.CheckOnClick = $true #so we only have to click once to check a box
    foreach ($Sku in Get-AzureADSubscribedSku | Select-Object -Property Sku*,ConsumedUnits -ExpandProperty PrepaidUnits) {
        Clear-Variable HRSku -ErrorAction SilentlyContinue
        $HRSku = $SkuToFriendly.Item("$($Sku.SkuID)")
        if ($HRSku){ $CheckedListBoxOutput = $HRSku + " -- " + ($Sku.Enabled-$Sku.ConsumedUnits) + " of " + $Sku.Enabled + " Available"}
        else { $CheckedListBoxOutput = $Sku.SkuPartNumber + " -- " + ($Sku.Enabled-$Sku.ConsumedUnits) + " of " + $Sku.Enabled + " Available"}
        $null = $CheckedListBox.Items.Add($CheckedListBoxOutput)
    }
    $CheckedListBox.ClearSelected()
    $CheckedListBox.Dock = [System.Windows.Forms.DockStyle]::Fill
    $LicenseSelectWindow.Controls.Add($CheckedListBox)

    $OKButton = New-Object System.Windows.Forms.Button
    $OKButton.Text = "Use Selected Licenses"
    $OKButton.Dock = [System.Windows.Forms.DockStyle]::Bottom
    $OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $LicenseSelectWindow.Controls.Add($OKButton)

    #put form in front of other windows
    $LicenseSelectWindow.TopMost = $true

    #display the form
    $null = $LicenseSelectWindow.ShowDialog()
    #####End License Selection Form#####

    #####License Selection Functionality#####
    if ($OKButton.DialogResult -eq "OK") {
        foreach($checkedlicense in $CheckedListBox.CheckedItems){
            Clear-Variable converttosku -ErrorAction SilentlyContinue
            Clear-Variable License -ErrorAction SilentlyContinue
            Clear-Variable LicensesToAssign -ErrorAction SilentlyContinue
            $converttosku = $checkedlicense -replace '\s--\s.*'
            $License = New-Object -TypeName Microsoft.Open.AzureAD.Model.AssignedLicense
            $License.SkuID = $FriendlyToSku.Item("$($converttosku)")
            $LicensesToAssign = New-Object -TypeName Microsoft.Open.AzureAD.Model.AssignedLicenses
            $LicensesToAssign.AddLicenses = $License
            Set-AzureADUserLicense -ObjectId $UPN -AssignedLicenses $LicensesToAssign
        }
    }
        
    #####Group Out-GridView to Select Groups#####
    # Pull User ObjectID and Group ObjectID to add member to all groups selected, skipping dynamic
    Clear-Variable user -ErrorAction SilentlyContinue
    Clear-Variable group -ErrorAction SilentlyContinue
    $user = Get-AzureADUser -ObjectID $UPN
    MailboxExistCheck
    Write-Verbose "Mailbox Exists, Please Select Groups To Add"
    foreach($group in Get-AzureADMSGroup | Where-Object {$_.GroupTypes -notcontains "DynamicMembership"} | Select-Object DisplayName,Description,ObjectId | Sort-Object DisplayName | Out-GridView -Passthru -Title "Hold Ctrl to select multiple groups" | Select-Object -ExpandProperty ObjectId){
        Add-AzureADGroupMember -ObjectId $group -RefObjectId $user.ObjectID
    }
    Write-Verbose "Selected Groups Added"
    
#Create Quit Prompt and Close While Loop
$quitboxOutput = [System.Windows.Forms.MessageBox]::Show("Do you need to create another user?" , "User Creation Complete" , 4)
}