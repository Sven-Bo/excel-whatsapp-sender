Attribute VB_Name = "mGlobals"
Option Explicit

''' *************************************************************************
''' Global Constant Declarations Follow
''' *************************************************************************

''' Version
Public Const VERSION_NUMBER As String = "3.5"


''' Only relevant for PRO Version
Public Const LINK_SUPPORT_FORM As String = "https://pythonandvba.com/whatsapp-pro-support"
Public Const LINK_RELEASE_NOTES As String = "https://pythonandvba.com/whatsapp-pro-releases"
Public Const LINK_FEATURE_REQUEST As String = "https://pythonandvba.com/go/whatsappblaster-feedback"
Public Const LINK_USAGE_GUIDELINES As String = "https://pythonandvba.com/go/whatsappblaster-usage-guidelines"
Public Const TUTORIAL_PHONE_NUMBER_CONVERTER_LINK As String = "https://pythonandvba.com/go/whatsappblaster-phone-number-converter-tutorial"
Public Const TUTORIAL_PHONE_NUMBER_REPLICATOR_LINK As String = "https://pythonandvba.com/go/whatsappblaster-phone-number-replicator-tutorial"
Public Const PLACEHOLDER_TUTORIAL_LINK As String = "https://pythonandvba.com/go/whatsappblaster-placeholder-tutorial"
Public Const PLACEHOLDER_TABLE_NAME = "tbl_PLACEHOLDER_DATA"


''' Chrome Driver Settings
Public Const ImplicitWait As Long = 9000          'Default 3000
Public Const PageLoad As Long = 90000             'Default 60000
Public Const TimeoutServer As Long = 90000        'Default 90000

''' First Row Used For The Iteration To Send Message
Public Const FirstRow As Integer = 4

''' Columns in worksheets
Public Enum BotColumn
    wcNumber = 2                                  ' col B
    wcText = 3                                    ' col C
    wcStatus = 4                                  ' col D
End Enum

''' Error Message
Public Const ERROR_EMAIL As String = "sven@pythonandvba.com"

''' Redirection Links
Public Const LINK_FAQ As String = "https://pythonandvba.com/whatsapp-faq"
Public Const LINK_MESSAGE_DETAILS_DOCS As String = "https://pythonandvba.com/go/whatsappblaster-message-details-documentation"
Public Const LINK_WHATSAPP_PRO As String = "https://pythonandvba.com/go/whatsapp-pro-purchase"

