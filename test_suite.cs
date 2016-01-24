/*
 * Connor Sheehan
 * sheehacg@mcmaster.ca
 * 
 * This program is the exe for the UOIT Software Deployment test suite.
 */


using System;
using System.Collections.Generic;
using Vestris.VMWareLib;
using System.Data.SqlClient;
using System.Text.RegularExpressions;
using System.IO;
using MySql.Data.MySqlClient;
using System.Threading;
using ImageMagick;

namespace TestSuite
{
    class Program
    {
        /*
         * Strings
         * *******
         * compId       - Computer ID on LANDesk
         * domainName   - domain to test on
         * hostName     - name of host
         * hostUser     - username of host user@domain.ca
         * hostPass     - password of hostUser
         * imageName    - name of image running tests on
         * picDir       - directory to save pics to for current test
         * psswd        - password for test user
         * pstempGuest  - location of pstemp on guest OS
         * pstempHost   - location of pstemp on Host
         * resultDir    - result directory for during-deployment screenshots
         * screenShotsDir
         *              - location of screenshot share
         * sfwrString   - ConnectionString for MySql database
         * snapName     - name of rollback snapshot
         * taskString   - ConnectionString for tasks database
         * pstempGuest  - location of pstemp dir on guest
         * usrnm        - username of test user
         * vmxPath      - path to vmx file for testing
         * workingDir   - location of TestSuite.exe on call to main()
         */
        private static string compId;
        private static string domainName;
        private static string hostName;
        private static string hostUser;
        private static string hostPass;
        private static string imageName;
        private static string picDir;
        private static string psswd;
        private static string pstempGuest = @"C:\pstemp";
        private static string pstempHost;
        private static string resultDir;
        private static string screenShotsDir;
        private static string sfwrString;
        private static string snapName = "deployPoint";
        private static string taskString;
        private static string usrnm;
        private static string vmxPath;
        private static string workingDir;



        /*
         * Integers
         * ********
         * count        - number of deployments tested on current snapshot
         * picTime      - how ofter to take a pic, in milliseconds
         * testPerSnap  - number of deployments to test before rolling back
         */
        private static int count = 0;
        private static int picTime = 3000;
        private static int testPerSnap;




        /*
         * Booleans
         * ********
         * useServer    - true? use vsphere : use workstation
         */
        private static bool? useServer;
        private static bool consoleOut = true;




        /*
         * Objects
         * *******
         * idPattern            - Pattern to map taskID to policy file
         * policyDirectoryFiles - list of policy files
         * tests                - Dictionary of tests, with taskID as key
         * conn                 - sql connection object
         * cmd                  - sql command object
         * reader               - sql reader object
         * myCmd                - mySql command object
         * myDeleteCommand      - mySql delete command object
         * myConn               - mySql connection object
         * myReader             - mySql reader object
         * stopwatch            - stopwatch object for timing deployments
         * pngs                 - collection of pictures for conversion into gif format
         */
        private static Regex idPattern = new Regex(@"CP.\d{4}.");
        private static List<string> policyDirectoryFiles;
        private static Dictionary<int, TestPackage> tests;
        private static MySqlCommand myCmd, myDeleteCmd;
        private static MySqlConnection myConn;
        private static MySqlDataReader myReader;
        private static SqlConnection conn;
        private static SqlCommand cmd;
        private static SqlDataReader reader;
        private static System.Diagnostics.Stopwatch stopwatch = new System.Diagnostics.Stopwatch();
        private static VMWareSnapshot deployPoint;



        // Main method. Start point of the program
        static void Main(string[] args)
        {
            /*
             * Print a banner to the screen
             */
            Console.WriteLine(Properties.Resources.banner);



            /*
             * Do nothing if no arguments given.
             * 
             * if the argument is -f, output to a file instead of console.
             */
            if (args.Length == 0 || args == null) Environment.Exit(0);
            else if (args[0] == "-f") consoleOut = false;



            /*
             * Get the current directory and store in a variable to avoid multiple calls
             */
            workingDir = Directory.GetCurrentDirectory();



            /*
             * Check to ensure that the pstemp directory is present in the same directory as
             * the TestSuite.exe file, with all important files included
             */
            pstempHost = workingDir + @"\pstemp";
            if (!Directory.Exists(pstempHost))
            {
                Console.WriteLine("[!] Error: the pstemp directory does not exist.");
                Environment.Exit(0);
            }
            else if (!File.Exists(pstempHost + @"\hstart64.exe") ||
                    !File.Exists(pstempHost + @"\PsExec.exe"))
            {
                Console.WriteLine("[!] Error: one or more files missing from pstemp directory.");
                Environment.Exit(0);
            }




            /*
             * Read lines from teh variables.ini file and set variables accordingly
             */
            string[] vars = File.ReadAllLines(workingDir + @"\Resources\variables.ini");
            foreach (string var in vars)
            {
                if (var.Contains("DOMAIN=")) domainName = var.Replace("DOMAIN=", "");
                else if (var.Contains("USERNAME=")) usrnm = var.Replace("USERNAME=", "");
                else if (var.Contains("PASSWORD=")) psswd = var.Replace("PASSWORD=", "");
                else if (var.Contains("VMX=")) vmxPath = var.Replace("VMX=", "");
                else if (var.Contains("CPUID=")) compId = var.Replace("CPUID=", "");
                else if (var.Contains("TESTSPERSNAP=")) testPerSnap = int.Parse(var.Replace("TESTSPERSNAP=", ""));
                else if (var.Contains("TASKCNXNSTRING=")) taskString = var.Replace("TASKCNXNSTRING=", "");
                else if (var.Contains("IMAGENAME=")) imageName = var.Replace("IMAGENAME=", "");
                else if (var.Contains("PICTIME=")) picTime = int.Parse(var.Replace("PICTIME=", ""));
                else if (var.Contains("USESERVER=")) useServer = bool.Parse(var.Replace("USESERVER=", ""));
                else if (var.Contains("HOSTNAME=")) hostName = var.Replace("HOSTNAME=", "");
                else if (var.Contains("HOSTUSER=")) hostUser = var.Replace("HOSTUSER=", "");
                else if (var.Contains("HOSTPASS=")) hostPass = var.Replace("HOSTPASS=", "");
                else if (var.Contains("SFWRCNXNSTRING=")) sfwrString = var.Replace("SFWRCNXNSTRING=", "");
                else if (var.Contains("SCREENSHOTSDIR=")) screenShotsDir = var.Replace("SCREENSHOTSDIR=", "");
            }





            /*
             * If some variable is not initialized, exit.
             */
            if (domainName == null ||
                usrnm == null ||
                psswd == null ||
                vmxPath == null ||
                compId == null ||
                taskString == null ||
                useServer == null ||
                sfwrString == null)
            {
                Console.WriteLine("[!] An essential variable has not been initiated. See variables.ini");
                Environment.Exit(0);
            }




            /*
             * If no image name was given, use the name of the VMX file
             */
            if (imageName == null) imageName = Path.GetFileNameWithoutExtension(vmxPath);
            screenShotsDir = screenShotsDir + imageName;



            /*
             * If the output is specified to file, redirect all Console.WriteLine to a file instead of the console.
             */
            if (!consoleOut)
            {
                FileStream fStream = new FileStream(screenShotsDir + @"\log\" + string.Format("{0:yyyy-MM-dd_hh-mm-ss-tt}", DateTime.Now) + ".txt", FileMode.OpenOrCreate, FileAccess.Write);
                StreamWriter fileOut = new StreamWriter(fStream);
                Console.SetOut(fileOut);
            }


            /*
             * If useServer is true and any of the essential server variables are empty, quit
             */
            if ((bool)useServer && (hostName == null ||
                            hostPass == null ||
                             hostUser == null))
            {
                Console.WriteLine("[!] Variables.ini is not configured for server use.");
                Environment.Exit(0);
            }




            /*
             * Create a new directory with the date and time stamp to store the test results
             */
            resultDir = workingDir + @"\Results\";
            System.IO.Directory.CreateDirectory(resultDir);


            resultDir = resultDir + imageName + @"\";
            System.IO.Directory.CreateDirectory(resultDir);






            /*
             * Create the autoLogon.reg file, used to allow auto-logon in the guest OS after snapshot revert

             * Sets to auto login
             * Turns off screensaver (broken?)
             * Turns off lock screen (broken?)
             */
            try
            {
                if (File.Exists(pstempHost + @"\autoLogon.reg")) File.Delete(pstempHost + @"\autoLogon.reg");
                string[] lines = { "Windows Registry Editor Version 5.00",
                                 @"[HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon]",
                                 "\"DefaultUserName\"=\"" + usrnm + "\"",
                                 "\"DefaultDomainName\"=\"" + domainName + "\"",
                                 "\"DefaultPassword\"=\"" + psswd + "\"",
                                 "\"AutoAdminLogon\"=\"1\"",
                                 "",
                                 @"[HKEY_CURRENT_USER\Control Panel\Desktop]",
                                 "\"ScreenSaveActive\"=\"0\"",
                                 "",
                                 @"[HKEY_LOCAL_MACHINE\SOFTWARE\Policies\Microsoft\Windows\Personalization]",
                                 "\"NoLockScreen\"=dword:00000001"};
                File.WriteAllLines(pstempHost + @"\autoLogon.reg", lines);
            }
            catch (Exception e)
            {
                Console.WriteLine("[!] Could not create auto-login registry file. You may need to do this manually.");
                Console.WriteLine("[?] Exception message: " + e.Message);
            }




            /*
             * Query the SQL database and use the reader to create a collection of 
             * the information
             */
            try
            {
                conn = new SqlConnection(taskString);
                conn.Open();
                cmd = new SqlCommand(Properties.Resources.taskQuery + compId, conn); // use the command resource, add the computer ID number
                reader = cmd.ExecuteReader();
                tests = new Dictionary<int, TestPackage>();
                while (reader.Read()) tests.Add(reader.GetInt32(1), new TestPackage(reader.GetString(0), reader.GetInt32(1).ToString()));
                reader.Close();
                conn.Close();
                Console.WriteLine("[i] " + tests.Count + " packages assigned to this machine.");

            }
            catch (SqlException sqle)
            {
                Console.WriteLine("Exception thrown: " + sqle.Message);
                Environment.Exit(0);
            }




            /*
             * Create VirtualHost and VirtualMachine objects with using directive
             */
            using (VMWareVirtualHost vhost = new VMWareVirtualHost())
            {


                /*
                 * Connect to either vSphere or Workstation, based on useServer variable
                 */
                if ((bool)useServer)
                {
                    vhost.ConnectToVMWareVIServer(hostName, hostUser, hostPass);
                    Console.WriteLine("[\u221A] Connected to vSphere.");
                }
                else
                {
                    vhost.ConnectToVMWareWorkstation();
                    Console.WriteLine("[\u221A] Connected to VMWareWorkstation.");
                }


                /*
                 * Open a virtual machine using the vmxPath
                 */
                using (VMWareVirtualMachine vm = vhost.Open(vmxPath))
                {
                    Console.WriteLine("[\u221A] Virtual machine opened.");




                    /*
                     * Power on the virtual machine and log in interactively.
                     * Requires user to log in interactively on the VM
                     */
                    try
                    {
                        vm.PowerOn();
                        Console.WriteLine("[\u221A] Powered on.");

                        vm.WaitForToolsInGuest();
                        Console.WriteLine("[\u221A] Tools found in guest.");


                        /*
                         * Login using the interactive login procedure
                         */
                        interactiveLogin(vm);



                        /*
                         * Copy necessary files for SYSTEM execution from host to guest.
                         * 
                         */
                        if (vm.DirectoryExistsInGuest(pstempGuest)) vm.DeleteDirectoryFromGuest(pstempGuest);
                        vm.CopyFileFromHostToGuest(pstempHost, pstempGuest);
                        Console.WriteLine("[\u221A] Files copied to temp directory " + pstempGuest + ".");





                     
                        /*
                         * Run the AutoLogon script for proper VM rollbacks
                         * We try try to delete the PSEXESVC process, if it exists. If it doesn't it throws an exception,
                         * which we ignore.
                         */
                        VMWareVirtualMachine.Process reg = detachSystemCommand(@"C:\Windows\regedit.exe", "/s \"" + pstempGuest + "\\autoLogon.reg\"", vm);
                        if (reg != null) while (vm.GuestProcesses.FindProcess(reg.Name, StringComparison.CurrentCulture) != null) ;
                        Console.WriteLine("[\u221A] Registry configured for auto logon.");




                        /*
                         * Create snapshot for deployment rollbacks
                         * If a snapshot 
                         */
                        try
                        {
                            deployPoint = vm.Snapshots.GetNamedSnapshot(snapName);
                            deployPoint.RemoveSnapshot();
                        }
                        catch (VMWareException) { }

                        deployPoint = vm.Snapshots.CreateSnapshot(snapName, "Start point for deployment tests");
                        Console.WriteLine("[\u221A] Snapshot 'deployPoint' captured.");


                        /*
                         * Sync policies with server.
                         */
                        policySync(vm);




                        /*
                         * Create a list of all files in the policy directory
                         * Remove all non-xml files from the list
                         * Look for the task ID in the file name, and map the file name to the correct package
                         */

                        policyDirectoryFiles = vm.ListDirectoryInGuest(@"C:\ProgramData\LANDesk\Policies", false);
                        policyDirectoryFiles.RemoveAll(x => !(x.Contains(".xml")));

                        foreach (string dir in policyDirectoryFiles)
                        {
                            Match m = idPattern.Match(dir);
                            int key = (m.Success) ? Int32.Parse(m.Value.Substring(3, 4)) : -1;
                            if (tests.ContainsKey(key)) tests[key].Cmd = dir;
                        }
                        Console.WriteLine((policyDirectoryFiles.Count == tests.Count) ? "[\u221A] All packages can be deployed." : "[!] There are " + policyDirectoryFiles.Count + " policy files and " + tests.Count + " tasks assigned.");





                        /*
                         * Rrun through each entry in the dictionary of test cases,
                         * 
                         */
                        using (myConn = new MySqlConnection(sfwrString))
                        {
                            myConn.Open();
                            foreach (KeyValuePair<int, TestPackage> entry in tests)
                            {
                                TestPackage test = entry.Value;

                            ReTest:
                                try
                                {


                                    /*
                                     * Query SLM, if any result is given test is a pass else fail
                                     */
                                    myCmd = new MySqlCommand("SELECT * FROM software WHERE task_id=" + test.TaskID, myConn);
                                    myReader = myCmd.ExecuteReader();
                                    test.NewResult = new TestResult("SLM Entry Exists", (myReader.HasRows) ? "1" : "0", "");
                                    myReader.Close();



                                    /*
                                     * If there is no xml file for the package, continue
                                     * For the package that cannot be deployed, add a test result saying so
                                     * If it can be deployed, add a positive test result
                                     */
                                    if (!test.isDeployable)
                                    {

                                        Console.WriteLine("[!] Continuing, cannot deploy " + test.Name);
                                        test.NewResult = new TestResult("LANDesk Task Available", "0", "Policy file not found.");
                                        continue;
                                    }
                                    else
                                    {
                                        Console.WriteLine("[*] Deploying " + test.Name);
                                        test.NewResult = new TestResult("LANDesk Task Available", "1", "");
                                    }



                                    /*
                                     * Directory for pictures resolved from resultDir and taskID
                                     */
                                    picDir = resultDir + test.TaskID + @"\";
                                    System.IO.Directory.CreateDirectory(picDir);




                                    /*
                                     * Get the number of registry entries in the LANDesk reg folder before deployment
                                     */
                                    int beforeDeploy = getRegEntryNum(vm);



                                    /*
                                     * Capture the process to determine when the deployment is complete
                                     * Start timing the deployment from when the PsExec process is detected
                                     */
                                    VMWareVirtualMachine.Process deployProcess = detachSystemCommand(@"C:\Program Files (x86)\LANDesk\LDClient\SDCLIENT.EXE", test.Cmd, vm);
                                    stopwatch.Restart();




                                    /*
                                     * While the program deploys
                                     * if 10000ms has passed since last pic taken, take a pic and increment
                                     * if the test is taking too long, add a failed test, revert snapshot, convert
                                     * png files to single gif and go to the next test
                                     */
                                    int pictureCount = 0;
                                    while (vm.GuestProcesses.FindProcess(deployProcess.Name, StringComparison.CurrentCulture) != null)
                                    {
                                        if (stopwatch.ElapsedMilliseconds > pictureCount * picTime) vm.CaptureScreenImage().Save(picDir + pictureCount++.ToString() + @".png");
                                        if (stopwatch.ElapsedMilliseconds > test.MaxTime)
                                        {
                                            test.NewResult = new TestResult("LANDesk Task Completed", "0", "");
                                            test.Pngs = Directory.GetFiles(picDir);
                                            test.PicsDir = picDir;
                                            revertAndLogin(vm);
                                            goto NextTest;
                                        }
                                    }
                                    stopwatch.Stop();
                                    test.Time = stopwatch.ElapsedMilliseconds;
                                    test.NewResult = new TestResult("LANDesk Task Completed", "1", test.Time.ToString());
                                    detachSystemCommand(@"C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe", "-command \"& {&\'Stop-Process\' -processname htmldialog}\"", vm);
                                    Console.WriteLine("[\u221A] Task complete. Task took " + (test.Time / 1000) + " seconds.");





                                    /*
                                     * The task has successfully completed running, we now add to the TestPackage object the
                                     * names of all .png files and the location of these files, for later gif creation.
                                     */
                                    test.Pngs = Directory.GetFiles(picDir);
                                    test.PicsDir = picDir;





                                    /*
                                     * Test if a new registry key was added. If there are more registry entries now than before the deployment,
                                     * pass. Otherwise fail.
                                     */
                                    test.NewResult = new TestResult("Registry Key Added", (beforeDeploy < getRegEntryNum(vm)) ? "1" : "0", "");





                                    /*
                                     * Check to see if any of the exe's returned by query to database exist on the guest
                                     * 
                                     */
                                    bool anyExeOnGuest = false;
                                    myCmd = new MySqlCommand(Properties.Resources.exeQuery + test.TaskID, myConn);
                                    myReader = myCmd.ExecuteReader();
                                    if (!myReader.HasRows) anyExeOnGuest = true;
                                    else
                                    {
                                        while (myReader.Read())
                                        {
                                            string exeName = myReader.GetString(2);
                                            string fileId = myReader.GetString(0);
                                            if (vm.FileExistsInGuest(exeName))
                                            {
                                                /*
                                                 * If the file exists on the guest; run the exe, wait 60 seconds and take a pic.
                                                 */
                                                anyExeOnGuest = true;
                                                VMWareVirtualMachine.Process testExe = vm.DetachProgramInGuest(myReader.GetString(2));
                                                stopwatch.Restart();
                                                while (stopwatch.ElapsedMilliseconds < 60000 && vm.GuestProcesses.FindProcess(testExe.Name, StringComparison.CurrentCulture) != null) ;
                                                if (vm.GuestProcesses.FindProcess(testExe.Name, StringComparison.CurrentCulture) != null)
                                                {
                                                    vm.CaptureScreenImage().Save(picDir + "exescreen_" + fileId + ".png");
                                                    try
                                                    {
                                                        testExe.KillProcessInGuest();
                                                    }
                                                    catch { }
                                                }
                                                detachSystemCommand(@"C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe", "-command \"& {&\'Stop-Process\' -processname @(javaw, java)}\"", vm);
                                                test.NewResult = new TestResult("Ran EXE - " + Path.GetFileName(exeName), "1", fileId);
                                            }
                                        }// end myReader.Read() while
                                    }





                                    /*
                                     * If there are no exe's on the guest, result is a fail.
                                     * If there are, result is a pass with optional message
                                     */
                                    test.NewResult = (!anyExeOnGuest) ? new TestResult("Files Installed", "0", "") : new TestResult("Files Installed", "1", (!myReader.HasRows) ? "No executables returned" : "");
                                    myReader.Close();





                                    /*
                                     * 
                                     * If the number of tests completed is greater than the revert
                                     * threshold, revert and continue testing.
                                     * Otherwise, kill the htmldialog.exe process for the next test.
                                     */
                                    if (++count == testPerSnap) revertAndLogin(vm);
                                }
                                catch (Exception e)
                                {
                                    /*
                                     * If this code is reached, test has not completed.
                                     * If this is not the first test deployed since a test passed, revert and try the test again.
                                     * Otherwise, revert and move to the next test.
                                     * Add a test result containing some info from Exception e
                                     */
                                    if (count != 0)
                                    {
                                        Console.WriteLine("[!] Test suite error.");
                                        Console.WriteLine("[!] Restarting test of " + test.Name);
                                        revertAndLogin(vm);
                                        goto ReTest;
                                    }
                                    test.NewResult = new TestResult("Test Suite Error", "0", e.Message);
                                    Console.WriteLine("[!] Exception caught, rolling back then moving to next test: ");
                                    Console.WriteLine("Type: " + e.GetType().ToString());
                                    Console.WriteLine("Message: " + e.Message);
                                    Console.WriteLine("Stack trace:\n" + e.StackTrace);
                                    revertAndLogin(vm);
                                }

                            NextTest:;
                            }// end foreach


                            try
                            {
                                /*
                             * We now have a dictionary of all the tests with taskID as the keys.
                             * Send the results to the SQL server
                             */
                                Console.WriteLine("[\u221A] Sending test results to database.");
                                myDeleteCmd = new MySqlCommand("DELETE FROM unit_test WHERE image=?image AND task_id=?task");
                                myDeleteCmd.Parameters.AddWithValue("?image", imageName);
                                myDeleteCmd.Parameters.AddWithValue("?task", "");
                                myDeleteCmd.Connection = myConn;
                                myDeleteCmd.CommandType = System.Data.CommandType.Text;

                                myCmd = new MySqlCommand(@"INSERT INTO unit_test (task_id, image, test_name, result, message) VALUES (?task_id, ?image, ?test_name, ?result, ?message)");
                                myCmd.CommandType = System.Data.CommandType.Text;
                                myCmd.Connection = myConn;
                                myCmd.Parameters.AddWithValue("?image", imageName);
                                myCmd.Parameters.AddWithValue("?task_id", "");
                                myCmd.Parameters.AddWithValue("?test_name", "");
                                myCmd.Parameters.AddWithValue("?result", "");
                                myCmd.Parameters.AddWithValue("?message", "");

                                foreach (KeyValuePair<int, TestPackage> entry in tests)
                                {

                                    myDeleteCmd.Parameters["?task"].Value = entry.Value.TaskID;
                                    myDeleteCmd.ExecuteNonQuery();

                                    myCmd.Parameters["?task_id"].Value = entry.Value.TaskID;
                                    foreach (TestResult result in entry.Value.AllResults)
                                    {
                                        myCmd.Parameters["?test_name"].Value = result.Name;
                                        myCmd.Parameters["?result"].Value = result.Result;
                                        myCmd.Parameters["?message"].Value = result.Message;
                                        myCmd.ExecuteNonQuery();
                                    }
                                }

                                Console.WriteLine("[\u221A] Results committed to database.");
                            }


                            /*
                             * If the database commit failed, print all results to a file in the working directory.
                             */
                            catch (Exception)
                            {
                                string fileName = screenShotsDir + @"\log\" + string.Format("{0:yyyy-MM-dd_hh-mm-ss-tt}", DateTime.Now) + ".txt";
                                FileStream fStream = new FileStream(fileName, FileMode.OpenOrCreate, FileAccess.Write);
                                StreamWriter fileOut = new StreamWriter(fStream);
                                foreach(KeyValuePair<int, TestPackage> pair in tests)
                                {
                                    foreach(TestResult result in pair.Value.AllResults)
                                    {
                                        fileOut.WriteLine(pair.Key.ToString() + "," + imageName + "," + result.Name + "," + result.Result + "," + result.Message);
                                    }
                                }
                                Console.WriteLine("[\u221A] Results written to file.\n" + fileName);
                            }

                        }// end of myConn using statement




                        /*
                         * Gif creation
                         */
                        Console.WriteLine("[\u221A] Creating gifs.");
                        foreach (KeyValuePair<int, TestPackage> pair in tests) pair.Value.createGif();



                        /*
                         * Copy the screenshots and gifs to the share
                         */
                        CopyFolder(resultDir, screenShotsDir);

                        Console.WriteLine("[\u221A] Testing complete. Exiting.");
                    }// end try



                    /*
                     * If for whatever reason an exception is thrown, print the error message to the screen.
                     * Distinguish between VMWare exception and other exceptions.
                     */
                    catch (VMWareException vmwe)
                    {
                        Console.WriteLine("[!] VMware exception: " + vmwe.Message);
                        Console.WriteLine(vmwe.ToString());
                        Console.WriteLine(vmwe.Source);
                        Environment.Exit(0);
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine("[!] Exception: " + e.Message + "\n");
                        Console.WriteLine("[!] Type:" + e.GetType());
                        Console.WriteLine(e.Source);
                        Console.WriteLine(e.StackTrace);
                        Environment.Exit(0);
                    }
                }// End vm using
            }// End vhost using
        }// end main


        /*
         * File copy subroutine
         */
        private static void CopyFolder(string source, string dest)
        {

            if (!Directory.Exists(dest)) Directory.CreateDirectory(dest);
            string[] files = Directory.GetFiles(source);
            foreach (string file in files)
            {
                string name = Path.GetFileName(file);
                string destination = Path.Combine(dest, name);
                File.Copy(file, destination);
            }
            string[] folders = Directory.GetDirectories(source);
            foreach (string folder in folders)
            {
                string name = Path.GetFileName(folder);
                string destination = Path.Combine(dest, name);
                CopyFolder(folder, destination);
            }
        }




        /*
         * Rolls back the given virtual machine to the deployPoint snapshot,
         * powers the vm on and logs in to an interactive session as the standand user.
         * Then calls policySync
         */
        private static void revertAndLogin(VMWareVirtualMachine ivm)
        {
            Console.WriteLine("[i] Rolling back.");
            count = 0;
            deployPoint.RevertToSnapshot();

            ivm.PowerOn();
            ivm.WaitForToolsInGuest();
        Login:
            try { ivm.LoginInGuest(domainName + @"\" + usrnm, psswd, 8, 300); }
            catch (Exception) { goto Login; }
            Console.WriteLine("[\u221A] Logged in.");
            policySync(ivm);
        }






        /*
         * Detaches commands as the SYSTEM user in the guest OS
         * Added the /WAIT argument to have hstart64 run until it's called process (PsExec) is done. So tracking hstart is the most
         * effective way of figuring out if the process is alive.
         */
        private static VMWareVirtualMachine.Process detachSystemCommand(string command, string args, VMWareVirtualMachine sysVm)
        {
            VMWareVirtualMachine.Process tst, desiredProc = null, hstartWait = sysVm.DetachProgramInGuest(pstempGuest + @"\hstart64.exe", "/WAIT /UAC \"" + pstempGuest + "\\PsExec.exe /accepteula -i -s \"" + command + "\" " + args + "\"");
            while ((tst = sysVm.GuestProcesses.FindProcess(hstartWait.Name, StringComparison.CurrentCulture)) != null && desiredProc == null)
            {
                desiredProc = sysVm.GuestProcesses.FindProcess(Path.GetFileName(command), StringComparison.CurrentCulture);
            }
            return desiredProc;
        }



        /*
         * Procedure for logging in interactively
         */
        private static void interactiveLogin(VMWareVirtualMachine ivm)
        {
            try
            {
                Console.Write("[i] Log in as " + domainName + @"\" + usrnm + " on the virtual machine, then press enter in the console.");
                Console.ReadLine();
                ivm.LoginInGuest(domainName + @"\" + usrnm, psswd, 8, 300);
                Console.WriteLine("[\u221A] Logged in.");
            }
            catch (Exception)
            {
                Console.WriteLine("[!] You did not log in correctly. Wait for all scripts to run, then try again.");
                interactiveLogin(ivm);
            }
        }



        /*
         * Runs regedit export of LANDesk software tracking, pulls file from guest -> host and returns number of lines in the file.
         *
         */
        private static int getRegEntryNum(VMWareVirtualMachine ivm)
        {
            if (ivm.FileExistsInGuest(pstempGuest + @"\reg.txt")) ivm.DeleteFileFromGuest(pstempGuest + @"\reg.txt");
            VMWareVirtualMachine.Process proc = ivm.RunProgramInGuest(@"C:\Windows\System32\Reg.exe", "export \"HKEY_LOCAL_MACHINE\\SOFTWARE\\LANdesk\\SOFTWARE\" \"" + pstempGuest + "\\reg.txt\"");
            if (File.Exists(workingDir + @"\reg.txt")) File.Delete(workingDir + @"\reg.txt");
            ivm.CopyFileFromGuestToHost(pstempGuest + @"\reg.txt", workingDir + @"\reg.txt");
            return File.ReadAllLines(workingDir + @"\reg.txt").Length;
        }



        /*
        * Run PolicySync as system user
        * To wait for the completion of PolicySync, we look for PsExec.exe process on guest,
        * when null is returned the process does not exist
        */
        private static void policySync(VMWareVirtualMachine ivm)
        {
            VMWareVirtualMachine.Process psync = detachSystemCommand(@"C:\Program Files (x86)\LANDesk\LDClient\PolicySync.exe", "", ivm);
            if (psync != null) while (ivm.GuestProcesses.FindProcess(psync.Name, StringComparison.CurrentCulture) != null) System.Threading.Thread.Sleep(1000);
            Console.WriteLine("[\u221A] PolicySync complete.");
        }


    }// end class:Program


    //////////////////////////////////////////////////////////////////////////////////////////////////////////

    /*
     * The TestPackage class holds all relevant info about each deployment package to be tested
     * This includes all information gathered from the sql query, as well as all information gathered from 
     * forensics (ie did the package deploy) and screenshots taken during deployment
     */
    class TestPackage
    {
        /*
         * Fields for a test package
         * packageName          - name of the test deployment package in LANDesk
         * cmdLine              - policyfile command line switch
         * taskIDNum            - the LANDesk task ID for the package
         * pd                   - picture directory for gif creation
         * durationMillis       - length of install in milliseconds.
         * maxMillis            - max amount of time allowed for install. NOT IMPLEMENTED FOR USE
         * results              - list of test results for database
         * pngToGifFiles        - list of all png files to be added to gif
         */
        private string pd;
        private string packageName;
        private string cmdLine;
        private string taskIDNum;
        private long durationMillis;
        private long maxMillis;
        private List<TestResult> results;
        private string[] pngToGifFiles;


        /*
         * Properties of a test package
         * Only accessor allowed, except for Cmd which is set outside of the constructor
         */
        public string PicsDir { set { this.pd = value; } }
        public string[] Pngs { set { this.pngToGifFiles = value; } }
        public string Name { get { return this.packageName; } }
        public string TaskID { get { return this.taskIDNum; } }
        public long MaxTime { get { return this.maxMillis; } }
        public bool isDeployable { get { return this.cmdLine != null; } }
        public long Time
        {
            get { return this.durationMillis; }
            set { this.durationMillis = value; }
        }
        public string Cmd
        {
            get { return this.cmdLine; }

            // Value in this context should be the xml filename for the specified package
            set
            {
                this.cmdLine = "/policyfile=\"" + value + "\"";
            }
        }

        /*
         * Set new results and get all results
         */
        public TestResult NewResult { set { this.results.Add(value); } }
        public List<TestResult> AllResults { get { return this.results; } }


        /*
         * Create gif for this test packa4
         * If there are more pics than the max (250), set the skip value slightly higher to skip extra images
         */
        public void createGif()
        {
            if (this.pngToGifFiles == null || this.pd == null) return;
            Console.WriteLine("[i] Creating gif for task " + this.taskIDNum);
            using (MagickImageCollection pngs = new MagickImageCollection())
            {
                int skip = 1, maxPics = 250, extraPics;
                if ((extraPics = pngToGifFiles.Length - maxPics) > 1) skip = (int)Math.Ceiling((double)pngToGifFiles.Length / maxPics);
                for (int i = 0, count = 0; i < pngToGifFiles.Length; i += skip)
                {
                    pngs.Add(pngToGifFiles[i]);
                    pngs[count++].AnimationDelay = 10;
                }
                pngs[0].AnimationIterations = 1;
                QuantizeSettings sett = new QuantizeSettings();
                sett.Colors = 256;
                pngs.Quantize(sett);
                pngs.Optimize();
                pngs.Write(this.pd + "install_" + this.taskIDNum + ".gif");
                foreach (string fileToDelete in this.pngToGifFiles) File.Delete(fileToDelete); // Delete the extra .png files
                Console.WriteLine("[\u221A] Gif created for task " + this.taskIDNum + ".");
            } // end pngs using
        }



        /*
         * Constructor for a test package
         */
        public TestPackage(string pkn, string iD)
        {
            this.packageName = pkn;
            this.taskIDNum = iD;
            this.results = new List<TestResult>();
            this.maxMillis = 3600000;
        }
    }// end class:TestPackage



    /*
     * Record for a test Result
     */
    class TestResult
    {
        private string name;
        private string result;
        private string message;

        public string Name { get { return this.name; } }
        public int Result { get { return int.Parse(this.result); } }
        public string Message { get { return this.message; } }


        public TestResult(string n, string r, string m)
        {
            this.name = n;
            this.result = r;
            this.message = m;
        }
    }
}
