{0}------------------------------------------------

# **Build your first Project task pane add-in**

Article • 09/17/2024

In this article, you'll walk through the process of building a Project task pane add-in.

#### **Prerequisites**

- Node.js (the latest LTS version). Visit the [Node.js site](https://nodejs.org/) to download and install the right version for your operating system.
- The latest version of Yeoman and the Yeoman generator for Office Add-ins. To install these tools globally, run the following command via the command prompt.

command line npm install -g yo generator-office

7 **Note**

Even if you've previously installed the Yeoman generator, we recommend you update your package to the latest version from npm.

- Office connected to a Microsoft 365 subscription (including Office on the web).
#### 7 **Note**

If you don't already have Office, you might qualify for a Microsoft 365 E5 developer subscription through the **[Microsoft 365 Developer Program](https://aka.ms/m365devprogram)** ; for details, see the **[FAQ](https://learn.microsoft.com/en-us/office/developer-program/microsoft-365-developer-program-faq#who-qualifies-for-a-microsoft-365-e5-developer-subscription-)**. Alternatively, you can **[sign up for a 1-month free trial](https://www.microsoft.com/microsoft-365/try)** or **[purchase a Microsoft 365 plan](https://www.microsoft.com/microsoft-365/business/compare-all-microsoft-365-business-products-g)** .

- Project 2016 or later on Windows
## **Create the add-in**

Run the following command to create an add-in project using the Yeoman generator. A folder that contains the project will be added to the current directory.

{1}------------------------------------------------

#### 7 **Note**

When you run the yo office command, you may receive prompts about the data collection policies of Yeoman and the Office Add-in CLI tools. Use the information that's provided to respond to the prompts as you see fit.

When prompted, provide the following information to create your add-in project.

- **Choose a project type:** Office Add-in Task Pane project
- **Choose a script type:** JavaScript
- **What do you want to name your add-in?** My Office Add-in
- **Which Office client application would you like to support?** Project

After you complete the wizard, the generator creates the project and installs supporting Node components.

## **Explore the project**

The add-in project that you've created with the Yeoman generator contains sample code for a very basic task pane add-in.

- The **./manifest.xml** file in the root directory of the project defines the settings and capabilities of the add-in.
- The **./src/taskpane/taskpane.html** file contains the HTML markup for the task pane.

{2}------------------------------------------------

- The **./src/taskpane/taskpane.css** file contains the CSS that's applied to content in the task pane.
- The **./src/taskpane/taskpane.js** file contains the Office JavaScript API code that facilitates interaction between the task pane and the Office client application. In this quick start, the code sets the Name field and Notes field of the selected task of a project.

### **Try it out**

- 1. Navigate to the root folder of the project.
command line

cd "My Office Add-in"

- 2. Start the local web server.

```
7 Note
```
Office Add-ins should use HTTPS, not HTTP, even while you're developing. If you're prompted to install a certificate after you run one of the following commands, accept the prompt to install the certificate that the Yeoman generator provides. You may also have to run your command prompt or terminal as an administrator for the changes to be made.

Run the following command in the root directory of your project. When you run this command, the local web server will start.

command line

npm run dev-server

- 3. In Project, create a simple project plan.
- 4. Load your add-in in Project by following the instructions in Sideload Office Add-ins on Windows.
- 5. Select a single task within the project.
- 6. At the bottom of the task pane, choose the **Run** link to rename the selected task and add notes to the selected task.

{3}------------------------------------------------

|           | B                                                                                                                                                                     |      |                           | Simple project plan - Project Professional |                                             |        | Gantt Chart Tools      |                                         |                                                                      |                                                       |                 |   | D | ×      |
|-----------|-----------------------------------------------------------------------------------------------------------------------------------------------------------------------|------|---------------------------|--------------------------------------------|---------------------------------------------|--------|------------------------|-----------------------------------------|----------------------------------------------------------------------|-------------------------------------------------------|-----------------|---|---|--------|
|           | File                                                                                                                                                                  | lask | Resource                  | Report<br>Project                          | Help<br>View                                | Team   | Format                 |                                         |                                                                      | O Tell me what you want to do                         |                 | 3 | 0 | ×      |
|           | Subproject                                                                                                                                                            |      | Get Add-ins<br>My Add-ins | Project<br>Information Fields Sprints      | Custom Manage Links Between<br>Projects     |        | Change<br>Working Time | Calculate<br>Project Baseline · Project | Set<br>Move                                                          | Status Date: NA<br>= Update Project                   | ABC<br>Spelling |   |   |        |
|           | Insert                                                                                                                                                                |      | Add-ins                   |                                            | Properties                                  |        |                        | Schedule                                |                                                                      | Status                                                | Proofina        |   |   |        |
|           | New task name<br>d 4/24/19 - Fri 4/26/19<br>Today<br>Apr 28, '19<br>Start<br>Summary #1<br>Summary #2<br>Mon 4/22/19<br>Mon 4/29/19 - Thu 5/9/19<br>Mon 4/22/19 - Fri |      |                           |                                            | Sun 5/5/19                                  |        |                        |                                         | My Office Add-in                                                     |                                                       |                 |   |   | ×<br>> |
| IMELIN    |                                                                                                                                                                       |      |                           |                                            | May 5, "19<br>May 12, "19<br>Summary<br>Mon |        |                        | Finish<br>Wed 5/15/19                   |                                                                      | Discover what Office Add-ins<br>can do for you today! |                 |   |   |        |
|           |                                                                                                                                                                       |      | Task                      | Mode ▼ Task Name                           | 19<br>Duration                              | W<br>5 | Apr 21, '19<br>T       | Apr 28, "19<br>5<br>M                   |                                                                      | വ<br>Achieve more with Office integration             |                 |   |   |        |
|           | 2                                                                                                                                                                     |      | 1<br>P                    | 4 Summary #1<br>Task 1                     | 5 days<br>2 days                            |        |                        |                                         |                                                                      |                                                       |                 |   |   |        |
|           | 3                                                                                                                                                                     |      | 0                         | New task name                              | 3 days                                      |        |                        |                                         | Unlock features and functionality<br>Create and visualize like a pro |                                                       |                 |   |   |        |
|           |                                                                                                                                                                       |      | P                         | Summary #1 Complete                        | 0 days                                      |        |                        | 4/26                                    |                                                                      |                                                       |                 |   |   |        |
| ANTT CHAF | 5                                                                                                                                                                     |      | r                         | · Summary #2                               | 9 days                                      |        |                        |                                         |                                                                      |                                                       |                 |   |   |        |
|           | 6                                                                                                                                                                     |      | । ਤੇ                      | Task 3                                     | 3 days                                      |        |                        |                                         |                                                                      |                                                       |                 |   |   |        |
|           |                                                                                                                                                                       |      | P                         | Task 4                                     | 4 days                                      |        |                        |                                         |                                                                      | Modify the source files, then click Run.              |                 |   |   |        |
|           | 8                                                                                                                                                                     |      | P                         | Task 5                                     | 2 days                                      |        |                        |                                         |                                                                      |                                                       |                 |   |   |        |
|           | 9                                                                                                                                                                     |      | मी                        | Summary #2 Complete                        | 0 days                                      |        |                        |                                         |                                                                      |                                                       |                 |   |   |        |
|           | 10                                                                                                                                                                    |      | T                         | · Summary #3                               | 3 days                                      |        |                        |                                         |                                                                      |                                                       |                 |   |   |        |
|           |                                                                                                                                                                       |      |                           | Task 6                                     | 3 days                                      |        |                        |                                         |                                                                      |                                                       | Run             |   |   | >      |
| Ready     | ব                                                                                                                                                                     |      |                           | New Tasks : Manually Schecluled            | -4                                          |        |                        |                                         | -                                                                    |                                                       |                 |   |   |        |

- 7. When you want to stop the local web server and uninstall the add-in, follow these instructions:
	- To stop the server, run the following command.

| command line |  |  |  |  |  |
|--------------|--|--|--|--|--|
| npm stop     |  |  |  |  |  |

- To uninstall the sideloaded add-in, see Remove a sideloaded add-in.
#### **Next steps**

Congratulations, you've successfully created a Project task pane add-in! Next, learn more about the capabilities of a Project add-in and explore common scenarios.

**Project add-ins**

## **Troubleshooting**

- Ensure your environment is ready for Office development by following the instructions in Set up your development environment.
- Some of the sample code uses ES6 JavaScript. This isn't compatible with older versions of Office that use the Trident (Internet Explorer 11) browser engine. For information on how to support those platforms in your add-in, see Support older Microsoft webviews and Office versions. If you don't already have a Microsoft 365

{4}------------------------------------------------

subscription to use for development, you might qualify for a Microsoft 365 E5 developer subscription through the [Microsoft 365 Developer Program](https://aka.ms/m365devprogram) ; for details, see the [FAQ](https://learn.microsoft.com/en-us/office/developer-program/microsoft-365-developer-program-faq#who-qualifies-for-a-microsoft-365-e5-developer-subscription-). Alternatively, you can [sign up for a 1-month free trial](https://www.microsoft.com/microsoft-365/try) or [purchase a Microsoft 365 plan](https://www.microsoft.com/microsoft-365/business/compare-all-microsoft-365-business-products-g) .

- The automatic npm install step Yo Office performs may fail. If you see errors when trying to run npm start , navigate to the newly created project folder in a command prompt and manually run npm install . For more information about Yo Office, see Create Office Add-in projects using the Yeoman Generator.
#### **See also**

- Develop Office Add-ins
- Core concepts for Office Add-ins
- Using Visual Studio Code to publish