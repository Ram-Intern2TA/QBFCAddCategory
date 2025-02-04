using System;
using Interop.QBFC16; 

class QuickBooksDesktopCategory
{
    static void Main(string[] args)
    {
        try
        {
            // Create a new session manager object
            QBSessionManager sessionManager = new QBSessionManager();

            // Open a connection to QuickBooks
            sessionManager.OpenConnection("", "QuickBooks C# Application");

            // Begin a session to interact with QuickBooks
            sessionManager.BeginSession("", ENOpenMode.omDontCare);

            // Create the MsgSetRequest object to build the QBXML request
            IMsgSetRequest msgSetRequest = sessionManager.CreateMsgSetRequest("US", 16, 0); // US locale, version 16

            // Create an AccountAdd request within the MsgSetRequest
            IAccountAdd accountAddRequest = msgSetRequest.AppendAccountAddRq();

            // Set the properties for the new account (Category)
            accountAddRequest.Name.SetValue("Sample Category");  // Account (Category) name
            accountAddRequest.AccountType.SetValue(ENAccountType.atIncome);  // Account type (e.g., Income, Expense)

            // You can now use just AccountType, as AccountSubType might not be directly available
            // If you need more details on the type of account, you can use the AccountType property

            // Send the request to QuickBooks and get the response
            IMsgSetResponse msgSetResponse = sessionManager.DoRequests(msgSetRequest);

            // Process the response and print success/failure message
            IResponse response = msgSetResponse.ResponseList.GetAt(0);
            if (response.StatusCode == 0)  // Check if the request was successful
            {
                Console.WriteLine("Category added successfully.");
            }
            else
            {
                Console.WriteLine("Error: " + response.StatusMessage);
            }

            // End the session and close the connection
            sessionManager.EndSession();
            sessionManager.CloseConnection();
        }
        catch (Exception ex)
        {
            // Handle any errors that may occur
            Console.WriteLine("Error adding category: " + ex.Message);
        }
    }
}
