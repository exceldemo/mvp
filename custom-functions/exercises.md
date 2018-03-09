_[Workshop home](../index.md)_  >  _[Custom functions](index.md)_ > _Exercises_

# Exercises

## Make a code change

Modify the ADD42 function to subtract 42 instead:
  1. Click the customfunctions.js file in the online editor (or your own IDE) and search for the ADD42 function.
  2. Change the addition sign before the 42 to a subtraction sign.
  3. Go back to Excel and run the CONTOSO.RELOAD function to get your updates
  4. Try running ADD42 again to see if your change worked

## Change the visible name of the function:
  1. In customfunctions.js, search for the line of code where the "ADD42" string is declared in all caps, and change it to SUBTRACT42 to reflect the new correct behaviour.
  2. Reload your code and look for the change 

## Challenge #1: Change the add42 function to add 42 to just one number, not two.
Hint: you'll need to modify the function itself and the function metadata, all in the customfunctions.js file

## Challenge #2: Implement your own version of Excel's RAND() function called CONTOSO.RAND.
  - Hint: use the internet if you don't know how to do something in JavaScript
  - Note: we don't support volatile functions yet, so it won't recalculate unless you explicitly trigger it.

## Challenge #3: Write a new function called CONTOSO.LONGESTWORD that finds the longest word in a range
  - Hint: the secondHighestTemp function already looks at a range
  - Hint: careful about the data types!

## Challenge #4: Write your own version of Excel's VLOOKUP function
Note: we don't support optional parameters yet, so just assume they're all required.

## Challenge #5 (hard): The stock market is open right now! Write a function that streams the price of MSFT in real time
  - Hint: https://iextrading.com/developer/docs
  - Hint: https://stackoverflow.com/questions/247483/http-get-request-in-javascript


Thank you! Please take our [survey](https://forms.office.com/Pages/ResponsePage.aspx?id=v4j5cvGGr0GRqy180BHbR60_sgZbQMhNpsP2LBe2so9UMFVYNUFHVUZZUDRXT0czWkYxUzcyNDBYMy4u)


