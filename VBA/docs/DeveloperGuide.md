---
layout: page
title: Developer Guide
---
* Table of Contents
  {:toc}

--------------------------------------------------------------------------------------------------------------------

## **Acknowledgements**

* This project use only Microsoft Excel, and Excel VBA.
* This project does not use any external library. 

--------------------------------------------------------------------------------------------------------------------

## **Setting up, getting started**
:exclamation: **Caution:**  
Follow the steps in this guide precisely. Things may not work out if you deviate at any point.

1. **Clone the repository**
    - Fork this [repository](https://github.com/ExcelMagician/Automated-Rostering-System) to your own GitHub account, then clone the fork to your local machine using Git.

2. **Install prerequisites**
    - **Excel version:** Use Excel 2016 or later (or Microsoft 365) with VBA support.
    - **VBA environment:** No external SDKs are needed beyond what comes with Excel.

3. **Open the workbook and enable macros**
    - Launch Excel and open the main workbook (e.g., `Automated Rostering System.xlsm`).
    - If prompted, enable macros. Without macros enabled, the automated rostering features will not work.

4. **Import the VBA project (optional)**
    - If you wish to edit or inspect the VBA code, open the VBA editor (`Alt + F11`).
    - Make sure the project is referenced correctly. You can import/export modules if needed.

5. **Verify the setup**
    - Open the main roster workbook and run key macros (e.g., populating the roster, generating an analysis report).
    - Check that the macros complete without errors.
    - Confirm that sample data appears correctly in the “Analysis Report” sheets.

6. **Before writing code**
    - **Code style:** Use consistent indentation and meaningful module names for VBA.
    - **Macro security settings:** Ensure your organisation’s macro security settings allow the macros to run.



--------------------------------------------------------------------------------------------------------------------

## **Design**

### Architecture

The program is modular, with separate subs handling each duty type and functionality.

**Main components of the architecture**

The **`Main`** sub orchestrates the entire rostering process, calling other modules in sequence.
* Upon running, it clears the content of the current roster and populate the roster table with personnels by calling related subs.

The modules are categorised according to functionalities:

* [**`Assignment`**](): These modules containes subroutines that assign personnels to various duty types on the "Roster" sheet, respecting respective duty constraints.
* [**`PersonnelLists`**](): These modules manages operations on personnel list sheets, including inserting or deleting staff in tables.
* [**`Swap`**](): This module provides the functionality to swap duties between staff members on the roster.
* [**`Analysis`**](): These modules contain subroutines related to the generation of analysis reports based on roster data.
* [**`Utilities`**](): These modules provides subs for sheet protection, unprotection and other utility tasks.

[**`Helpers`**]() These modules contains utility subroutines that support other modules with common tasks

The sections below give more details of each category.

### Assignment

1) **AssignSatAOHDuties Module**

`AssignSatAOHDuties` sub: Iterates through the "Sat AOH PersonnelList" table and assigns available staffs to Saturday AOH duties across two columns on the "Roster" sheet. 

The IncrementDutiesCounter helper subroutine supports this by updating the "Duties Counter" column in the personnel table for each assigned staff member.

* Constraints
    
    * Staff are assigned only on rows where the day is "Sat" and the respective cell is empty
    * Assignments respect the maximum duty limit ("Max Duties") for each staff member, checked against their current duty count.
    * Consecutive Saturday assignments are prevented, ensuring no staff is assigned to the same duty two weeks in a row.
    * Two columns of Sat AOH duties must have two distinct staffs.

2) **AssignAOHDuties Module**

`AssignAOHDuties` sub: Iterates through the "AOH PersonnelList" table and the "AOHSpecificSaysWorkingStaff" table and assigns available staffs to AOH duties on the "Roster" sheet. 

Assignment logic:

1) First, it assigns staff with specific working days based on their defined schedules, shuffling eligible rows randomly to distribute duties evenly
2) Then, it assigns staff with "All Days" availability to the remaining slots.
3) Thirdly, it reassigns duties through a swapping mechanism if unfilled slots persists.
4) If reassignment still fails after 10 maximum tries, it fallbacks to assign remaining eligible staff to unfilled slots with yellow highlighting.

The IncrementDutiesCounter and DecrementDutiesCounter helper subroutine supports this by updating the "Duties Counter" column in the personnel table for each assigned staff member.

The CheckWeeklyLimit helper subroutine supports this by ensuring that all staff assigned to the AOH duties only work a maximum of 1 AOH duties per week.

* Constraints
    
    * Assignments are restricted to non-Saturday days with "SEM TIME" status and empty AOH slots.
    * For "specific working days staff", assignments are limited by their defined working days and overall "Max Duties", with a weekly limit of one duty per staff.
    * For "all days staff", assignments are constrained by their "Max Duties" and a weekly limit of one AOH duty per staff. Note that Sat AOH also counts as AOH duties.
    * Reassignment and swapping ensures no staff exceeds their weekly limit, and swaps require a week with no prior duties for the eligible staff, with a maxiumum of 10 reassignment attempts.
    * Fallback assignments highlight slots in yellow to indicate manual review may be needed.


### Logic component

**API** : [`Logic.java`](https://github.com/se-edu/addressbook-level3/tree/master/src/main/java/tassist/address/logic/Logic.java)

Here's a (partial) class diagram of the `Logic` component:

<img src="images/LogicClassDiagram.png" width="550"/>

The sequence diagram below illustrates the interactions within the `Logic` component, taking `execute("delete 1")` API call as an example.

![Interactions Inside the Logic Component for the `delete 1` Command](images/DeleteSequenceDiagram.png)

<div markdown="span" class="alert alert-info">:information_source: **Note:** The lifeline for `DeleteCommandParser` should end at the destroy marker (X) but due to a limitation of PlantUML, the lifeline continues till the end of diagram.
</div>

How the `Logic` component works:

1. When `Logic` is called upon to execute a command, it is passed to an `AddressBookParser` object which in turn creates a parser that matches the command (e.g., `DeleteCommandParser`) and uses it to parse the command.
1. This results in a `Command` object (more precisely, an object of one of its subclasses e.g., `DeleteCommand`) which is executed by the `LogicManager`.
1. The command can communicate with the `Model` when it is executed (e.g. to delete a person).<br>
   Note that although this is shown as a single step in the diagram above (for simplicity), in the code it can take several interactions (between the command object and the `Model`) to achieve.
1. The result of the command execution is encapsulated as a `CommandResult` object which is returned back from `Logic`.

Here are the other classes in `Logic` (omitted from the class diagram above) that are used for parsing a user command:

<img src="images/ParserClasses.png" width="600"/>

How the parsing works:
* When called upon to parse a user command, the `AddressBookParser` class creates an `XYZCommandParser` (`XYZ` is a placeholder for the specific command name e.g., `AddCommandParser`) which uses the other classes shown above to parse the user command and create a `XYZCommand` object (e.g., `AddCommand`) which the `AddressBookParser` returns back as a `Command` object.
* All `XYZCommandParser` classes (e.g., `AddCommandParser`, `DeleteCommandParser`, ...) inherit from the `Parser` interface so that they can be treated similarly where possible e.g, during testing.

### Model component
**API** : [`Model.java`](https://github.com/se-edu/addressbook-level3/tree/master/src/main/java/tassist/address/model/Model.java)

<img src="images/ModelClassDiagram.png" width="450" /><br>
Note that ***filtered** should be next to the arrow from `ModelManager` to `Person`.\
It is displayed incorrectly due to limitations of puml.

The `Model` component,

* manages the application's data through the `AddressBook` class, which contains:
    * A `UniquePersonList` for storing `Person` objects
    * A `UniqueTimedEventList` for storing `TimedEvent` objects
* stores the currently 'selected' `Person` objects (e.g., results of a search query) as a separate _filtered_ list which is exposed to outsiders as an unmodifiable `ObservableList<Person>` that can be 'observed' e.g. the UI can be bound to this list so that the UI automatically updates when the data in the list change.
* stores a `UserPref` object that represents the user's preferences. This is exposed to the outside as a `ReadOnlyUserPref` objects.
* does not depend on any of the other three components (as the `Model` represents data entities of the domain, they should make sense on their own without depending on other components)
* provides data manipulation operations:
    * CRUD operations for `Person` objects
    * CRUD operations for `TimedEvent` objects
    * Sorting operations using `Comparator<Person>` and `Comparator<TimedEvent>`
    * Filtering operations using `Predicate<Person>` and `Predicate<TimedEvent>`
* manages user preferences through:
    * `UserPrefs` for storing application settings
    * `GuiSettings` for UI-specific preferences
    * File path management for data persistence

The component follows the Observer pattern through JavaFX's `ObservableList` interface, allowing the UI to automatically update when the underlying data changes.

<div markdown="span" class="alert alert-info">:information_source: **Note:** An alternative (arguably, a more OOP) model is given below. It has a `Tag` list in the `AddressBook`, which `Person` references. This allows `AddressBook` to only require one `Tag` object per unique tag, instead of each `Person` needing their own `Tag` objects.<br>

<img src="images/BetterModelClassDiagram.png" width="450" />
The rest of the Person's attributes has been abstracted out in the image above.

While the current implementation does not use this alternative model for `Tag`, it does use this approach for `TimedEvent`.\
\
The `AddressBook` maintains a `UniqueTimedEventList` which enforces uniqueness between timed events using `TimedEvent#isSameTimedEvent(TimedEvent)`.\
\
This allows the `AddressBook` to only require one `TimedEvent` object per unique event (based on name and time), rather than each `Person` needing their own copy of the same event.

</div>

### Storage component

**API** : [`Storage.java`](https://github.com/se-edu/addressbook-level3/tree/master/src/main/java/tassist/address/storage/Storage.java)

<img src="images/StorageClassDiagram.png" width="550" />

The `Storage` component,
* can save both address book data and user preference data in JSON format, and read them back into corresponding objects.
* inherits from both `AddressBookStorage` and `UserPrefStorage`, which means it can be treated as either one (if only the functionality of only one is needed).
* depends on some classes in the `Model` component (because the `Storage` component's job is to save/retrieve objects that belong to the `Model`)

### Common classes

Classes used by multiple components are in the `tassist.address.commons` package.

--------------------------------------------------------------------------------------------------------------------

## **Implementation**

This section describes some noteworthy details on how certain features are implemented.

### \[Proposed\] Undo/redo feature

#### Proposed Implementation

The proposed undo/redo mechanism is facilitated by `VersionedAddressBook`. It extends `AddressBook` with an undo/redo history, stored internally as an `addressBookStateList` and `currentStatePointer`. Additionally, it implements the following operations:

* `VersionedAddressBook#commit()` — Saves the current address book state in its history.
* `VersionedAddressBook#undo()` — Restores the previous address book state from its history.
* `VersionedAddressBook#redo()` — Restores a previously undone address book state from its history.

These operations are exposed in the `Model` interface as `Model#commitAddressBook()`, `Model#undoAddressBook()` and `Model#redoAddressBook()` respectively.

Given below is an example usage scenario and how the undo/redo mechanism behaves at each step.

Step 1. The user launches the application for the first time. The `VersionedAddressBook` will be initialized with the initial address book state, and the `currentStatePointer` pointing to that single address book state.

![UndoRedoState0](images/UndoRedoState0.png)

Step 2. The user executes `delete 5` command to delete the 5th person in the address book. The `delete` command calls `Model#commitAddressBook()`, causing the modified state of the address book after the `delete 5` command executes to be saved in the `addressBookStateList`, and the `currentStatePointer` is shifted to the newly inserted address book state.

![UndoRedoState1](images/UndoRedoState1.png)

Step 3. The user executes `add n/David …​` to add a new person. The `add` command also calls `Model#commitAddressBook()`, causing another modified address book state to be saved into the `addressBookStateList`.

![UndoRedoState2](images/UndoRedoState2.png)

<div markdown="span" class="alert alert-info">:information_source: **Note:** If a command fails its execution, it will not call `Model#commitAddressBook()`, so the address book state will not be saved into the `addressBookStateList`.

</div>

Step 4. The user now decides that adding the person was a mistake, and decides to undo that action by executing the `undo` command. The `undo` command will call `Model#undoAddressBook()`, which will shift the `currentStatePointer` once to the left, pointing it to the previous address book state, and restores the address book to that state.

![UndoRedoState3](images/UndoRedoState3.png)

<div markdown="span" class="alert alert-info">:information_source: **Note:** If the `currentStatePointer` is at index 0, pointing to the initial AddressBook state, then there are no previous AddressBook states to restore. The `undo` command uses `Model#canUndoAddressBook()` to check if this is the case. If so, it will return an error to the user rather
than attempting to perform the undo.

</div>

The following sequence diagram shows how an undo operation goes through the `Logic` component:

![UndoSequenceDiagram](images/UndoSequenceDiagram-Logic.png)

<div markdown="span" class="alert alert-info">:information_source: **Note:** The lifeline for `UndoCommand` should end at the destroy marker (X) but due to a limitation of PlantUML, the lifeline reaches the end of diagram.

</div>

Similarly, how an undo operation goes through the `Model` component is shown below:

![UndoSequenceDiagram](images/UndoSequenceDiagram-Model.png)

The `redo` command does the opposite — it calls `Model#redoAddressBook()`, which shifts the `currentStatePointer` once to the right, pointing to the previously undone state, and restores the address book to that state.

<div markdown="span" class="alert alert-info">:information_source: **Note:** If the `currentStatePointer` is at index `addressBookStateList.size() - 1`, pointing to the latest address book state, then there are no undone AddressBook states to restore. The `redo` command uses `Model#canRedoAddressBook()` to check if this is the case. If so, it will return an error to the user rather than attempting to perform the redo.

</div>

Step 5. The user then decides to execute the command `list`. Commands that do not modify the address book, such as `list`, will usually not call `Model#commitAddressBook()`, `Model#undoAddressBook()` or `Model#redoAddressBook()`. Thus, the `addressBookStateList` remains unchanged.

![UndoRedoState4](images/UndoRedoState4.png)

Step 6. The user executes `clear`, which calls `Model#commitAddressBook()`. Since the `currentStatePointer` is not pointing at the end of the `addressBookStateList`, all address book states after the `currentStatePointer` will be purged. Reason: It no longer makes sense to redo the `add n/David …​` command. This is the behavior that most modern desktop applications follow.

![UndoRedoState5](images/UndoRedoState5.png)

The following activity diagram summarizes what happens when a user executes a new command:

<img src="images/CommitActivityDiagram.png" width="250" />

#### Design considerations:

**Aspect: How undo & redo executes:**

* **Alternative 1 (current choice):** Saves the entire address book.
    * Pros: Easy to implement.
    * Cons: May have performance issues in terms of memory usage.

* **Alternative 2:** Individual command knows how to undo/redo by
  itself.
    * Pros: Will use less memory (e.g. for `delete`, just save the person being deleted).
    * Cons: We must ensure that the implementation of each individual command are correct.

_{more aspects and alternatives to be added}_

### \[Proposed\] Data archiving

_{Explain here how the data archiving feature will be implemented}_


--------------------------------------------------------------------------------------------------------------------

## **Documentation, logging, testing, configuration, dev-ops**

* [Documentation guide](Documentation.md)
* [Testing guide](Testing.md)
* [Logging guide](Logging.md)
* [Configuration guide](Configuration.md)
* [DevOps guide](DevOps.md)

--------------------------------------------------------------------------------------------------------------------

## **Appendix: Requirements**

### Product scope

**Target user profile**:

This product is designed for staff and supervisors at the National University of Singapore (NUS) libraries who need to track and manage **duty rosters** efficiently. Such users typically:
* Manage medium to large teams of library staff and student helpers across multiple library branches.
* Need quick access to individual shift assignments, duties, and roles.
* Prefer a fast, mouse‑heavy workflows.
* Want automated scheduling to ensure fair distribution of duties based on office's restrictions and staff availability.
* Are comfortable using data‑entry forms and familiar with spreadsheet applications like Excel.
* Oversee multiple rosters spanning different months or semesters.

**Value proposition**:\
Provides an easy way for library supervisors and staff to manage shift assignments and generate rosters. 
The system automatically assigns duties based on office's restrictions and staff availability, generates reports comparing 
planned duties to actual duties, and summarises workloads across different shifts. This improves the fairness 
and efficiency of duty distribution while reducing administrative workload and the risk of manual errors.

### User stories

Priorities: High (must have) - `* * *`, Medium (nice to have) - `* *`, Low (unlikely to have) - `*`

| Priority | As a …​       | I want to …​                                                                                 | So that…​                                            |
|----------|---------------|----------------------------------------------------------------------------------------------|------------------------------------------------------|
| `* * *`  | supervisor    | automatically assign duites based on office's restrictions                                   | I can generate a fair and balanced roster quickly    |
| `* * *`  | staff member  | see my upcoming shifts                                                                       | I can plan my personal schedule in advance           |
| `* * *`  | staff member  | request a shift swap with a colleague                                                        | I can handle unexpected personal commitments         |
| `* * *`  | supervisor    | generate an analysis report that compares actual duties completed against the planned roster | I can understand workload distribution               |
| `* *`    | staff member  | indicate my non-working days                                                                 | I am not assigned shifts when I am unavailable       |
| `* *`    | supervisor    | protect the roster data with a password                                                      | only authorized users can modify the schedule        |
| `* *`    | administrator | add and remove staff members from the system                                                 | new hires and departures are reflected in the roster |
| `* *`    | supervisor    | make sure all the staff work at most one duties per day even after swapping                  | there are no violation on the office's restrictions  |
| `* *`    |               |                                                                                              |                                                      |
| `* *`    |               |                                                                                              |                                                      |
| `* *`    |               |                                                                                              |                                                      |
| `* *`    |               |                                                                                              |                                                      |
| `* *`    |               |                                                                                              |                                                      |
| `* *`    |               |                                                                                              |                                                      |
| `* *`    |               |                                                                                              |                                                      |
| `* *`    |               |                                                                                              |                                                      |
| `*`      |               |                                                                                              |                                                      |
| `*`      |               |                                                                                              |                                                      |
| `*`      |               |                                                                                              |                                                      |
| `*`      |               |                                                                                              |                                                      |
| `*`      |               |                                                                                              |                                                      |
| `*`      |               |                                                                                              |                                                      |
| `*`      |               |                                                                                              |                                                      |
| `*`      |               |                                                                                              |                                                      |
| `*`      |               |                                                                                              |                                                      |
| `*`      |               |                                                                                              |                                                      |
| `*`      |               |                                                                                              |                                                      |
| `*`      |               |                                                                                              |                                                      |

### Use cases

(For all use cases below, the **System** is `Automate Rostering System` and the **Actor** is the `user (supervisor)`, unless specified otherwise)

**Use case: UC1 - Add a staff into the personnel list**

**MSS**

1. User requests to add a staff member.
2. User enters the staff's details(e.g., name, department, working days if applicable).
3. System validates the input and adds the staff to the personnel list.
4. System confirms the successful addition.
5. Use case ends.

**Extensions**

* 3a. Invalid value for one or more input fields.
    * 3a1. System displays an error message indicating the invalid field.
    * 3a2. User re-enters the corrected value.
    * Use case resumes at step 3.

* 3b. Duplicate staff detected.
    * 3a1. System displays an error message indicating that the staff member already exists.
    * Use case ends.

*{More to be added}*

### Non-Functional Requirements

1. Should work on any mainstream OS with Microsoft Excel 2016 or later (or Microsoft 365 Desktop version).

2. Should be able to handle up to 300 staff entries and 6 months’ worth of shift data without noticeable slowdown.

3. A user with above-average Excel proficiency should be able to accomplish most scheduling and reporting tasks within a few clicks or macro activations.

4. Should generate shift reports and duty analyses within 5 seconds of macro execution.

5. Should store all data locally within Excel spreadsheets — no external database required.


*{More to be added}*

### Glossary

* **Mainstream OS**: Windows, Linux, Unix, MacOS
* **Private contact detail**: A contact detail that is not meant to be shared with others
* **TA (Teaching Assistant/Tutor)**: A university staff member who assists in teaching, grading,
  and managing students in a course.
* **CLI (Command Line Interface)**: A text-based interface that allows users to interact with the system
  using typed commands.

--------------------------------------------------------------------------------------------------------------------

## **Appendix: Instructions for manual testing**

Given below are instructions to test the app manually.

<div markdown="span" class="alert alert-info">:information_source: **Note:** These instructions only provide a starting point for testers to work on;
testers are expected to do more *exploratory* testing.

</div>

### Launch and shutdown

1. Initial launch

    1. Download the jar file and copy into an empty folder

    2. Double-click the jar file Expected: Shows the GUI with a set of sample contacts. The window size may not be optimum.

2. Saving window preferences

    1. Resize the window to an optimum size. Move the window to a different location. Close the window.

    2. Re-launch the app by double-clicking the jar file.<br>
       Expected: The most recent window size and location is retained.

3. _{ more test cases …​ }_

### Deleting a person

1. Deleting a person while all persons are being shown

    1. Prerequisites: List all persons using the `list` command. Multiple persons in the list.

    2. Test case: `delete 1`<br>
       Expected: First contact is deleted from the list. Details of the deleted contact shown in the status message. Timestamp in the status bar is updated.

    3. Test case: `delete 0`<br>
       Expected: No person is deleted. Error details shown in the status message. Status bar remains the same.

    4. Other incorrect delete commands to try: `delete`, `delete x`, `...` (where x is larger than the list size)<br>
       Expected: Similar to previous.

2. _{ more test cases …​ }_

--------------------------------------------------------------------------------------------------------------------

## **Planned Enhancement**
1. Users must provide inputs according to the parameters specified in the User Guide. If an invalid or unrecognized parameter is used, TAssist will treat it as an error related to the previous valid parameter.<br>
   For example: <br>
   `assignment n/quiz pr/30 d/22-11-2027`<br>
   `assignment n/quiz ab/xx d/22-11-2027`<br>
   Both examples will result in an error message related to the `name` parameter:<br>
   "Name should only contain alphanumeric characters and spaces, and it should not be blank."<br>
   This is because pr/30 and ab/xx are not valid parameters for the assignment command.
   This behavior will be improved in future versions of TAssist to provide more specific error messages.

2. **When using multiple screens**, if you move the application to a secondary screen, and later switch to using only the primary screen, the GUI will open off-screen. The remedy is to delete the `preferences.json` file created by the application before running the application again.
3. Enhance the `unassign` command to support unassigning TimedEvents from individual students
4. Add more preferences to `preferences.json` such as the maintaining the theme and the ratio of student display and the command area.
