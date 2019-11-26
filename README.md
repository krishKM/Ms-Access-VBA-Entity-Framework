# Ms-Access-VBA-Entity-Framework
An attempt to move away from plain SQL executions in Ms Access environment. Bringing LINQ / EF like operations for MS Access data tables. However, this project is only beneficial to those who do SQL operations in VBA.


We all agree that, due to the nature of `VBA`, it's quite easy to write a complex/clumsy code :<br/>
```VBA
  'Select all apples from fruit table
   Set Rs = currentDb.openRecordset("SELECT * FROM fruits where typeid = " & appleId)
   
   'Select fruit name from the result set
   fruitName = rs("fruitName")
   
   'Update all apple price to 100
   currentDb.Execute "Update fruits SET price = 100 where typeId = " & appleId)
```
Nothing wrong with the above code but what really bothers me is, that I have to remember the table names and field names everytime I do a SQL operation. Plus in case if a field-name has changed, a global search and replace is required.

Besides, I really want to get rid of joining `strings` like
```VBA
  updateCommand = "UPDATE fruits SET " & _
  "price =" & price & "," & _
  "description='" & description & "', " & _
  "updated_by ='" & staffName & "' " & _
  "where fruitId = " & fruitId
  
  currentDb.Execute updateCommand, dbFailOnError
```

See how quickly the `updateCommand` became clumsy? plus the code is now vulnerable to `SQL injection` attacks too. Surely there are ways to prevent `sql injections` in VBA like using creting `query objects` with parameters, or custom functions that takes care of SQL injection. Most of them still having to build an SQL command.

<HR>

In this project I'm trying to achieve something like this: 
```VBA
  Dim fruits as new TFruits
  'Count number of furits from fruits table that are apples
  totalApples = fruits.where(typeId = appleId).count
  
  'Show the first fruit name from the fruit table
  me.txtFruitName = fruits.FirstOrDefault().FruitName ' Apple

  'Update fruit table and set the quantity to 60 for apples
  fruits.where(typeId = appleId)
  fruits.quantity = 60
  fruits.updated_by = staffName
  fruits.Update
  
```
<hr>

### Aim of this project
Pretty simple. Prevent or reduce writing plain SQL codes or having to remember table/field names when working with objects. Would like to use `intelisense` when selecting a property like `furits.fruitName`.

### Why?
 
1> Because it's much easier and cleaner to read the code.<br/>
2> If a table has to be replaced, renamed or any other operation, do it in one place.<br/>
3> reduce ugly SQL and string joins when building SQL commands.<br/>
4> <i>Feel free to add your answers here...</i>


### How?
The fun part begins here, by design, in Access VBA, it's not possible to treat a table as a class. For that reason, we need to generate a `class` from a table. Suppose we have a table called fruits like this:
```
Fruits
+---------+-------------+--------+-------+----------------+
| fruitId |  fruitName  | typeId | price |  description   |
+---------+-------------+--------+-------+----------------+
|       1 | Red Apple   |      1 | £0.40 | From Australia |
|       2 | Green Apple |      1 | £0.40 | From Australia |
|       3 | Banana      |      2 | £0.20 | Spain          |
|       4 | Pink Lady   |      1 | £0.40 | USA            |
+---------+-------------+--------+-------+----------------+

FruitTypes
+--------+-----------+
| typeId |   type    |
+--------+-----------+
|      1 | Apple     |
|      2 | Banana    |
|      3 | Guava     |
|      4 | Pineapple |
|      5 | Grape     |
+--------+-----------+


```
we need to make a `class` for Fruits that has `fruitId, fruitName, typeId, price, description` properties. -<i>Good old `Getters and setters` we've learned in the school</i>-.<br/>
One problem though, doing this manually is just boring and time consuming. So we are going to make a script that can generate class out of tables. Similar to EF does it if you go for `Database first` method.

Notice, the `Fruits` table has a `typeId`. TypeId is a `foreignKey` pointing to `FruitTypes` table. That means we should be able to do something like this if we succeed with this project.
```VBA
  Dim fruits      As new TFruits
  Dim fruitTypes  As new TFruitTypes
  Dim fruitType   As String
  
  'Method 1: Either join both fruits and fruitTypes and get the fruitType
  fruits.join JoinType.Inner, fruitTypes, fruits.typeId, fruitTypes.typeId
  fruits.where(fruits.typeId = appleId)
  fruitType = fruitTypes.FirstOrDefault().type ' => Apple
  
  'OR Method2: fruitTable should have fruitType property that returns the fruit type.
  fruitType = fruits.Where(fruits.typeId = appleId).FirstOrDefault().fruitTypes.type ' => Apple
  
```
Isn't that wonderful if we could use the tables this way? Clean code, no SQL strings?
(ᴗ_ ᴗ。)

Let's make a list of functions what this project should achieve. <br/>
1> Tables as classes<br/>
2> Tables can be joined<br/>
3> Standard `CRUD` operations like, creating new fruits, updating a fruit, delete a fruit etc.<br/>
4> Should be easy enough to use.<br/>
5> Scripts that can validate, generate, upgrade table structures

### Stay tuned for the end result..
