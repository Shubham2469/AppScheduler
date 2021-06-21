using System;
using System.Collections;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.ServiceProcess;
using System.Xml;
using System.Timers;
using System.Threading;
using System.Configuration;
using System.IO;

namespace AppScheduler
{
	public class AppScheduler : System.ServiceProcess.ServiceBase
	{
		string configPath;
		System.Timers.Timer _timer=new System.Timers.Timer();
		DataSet dsTasks=new DataSet();
		string formatString="MM/dd/yyyy HH:mm:ss";
		/// <summary> 
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

        /// <summary>
        /// Class that launches applications on demand.
        /// </summary>
		class AppLauncher
		{
			string app2Launch;
			public AppLauncher(string path)
			{
				app2Launch=path;
			}
			public void runApp()
			{
				ProcessStartInfo pInfo=new ProcessStartInfo(app2Launch);
				pInfo.WindowStyle=ProcessWindowStyle.Normal;
				Process p=Process.Start(pInfo);
			}
		}

		void timeElapsed(object sender, ElapsedEventArgs args)
		{
			DateTime currTime=DateTime.Now;
			foreach(DataRow dRow in dsTasks.Tables["task"].Rows)
			{
				DateTime runTime=Convert.ToDateTime(dRow["time"]);
				if(currTime>=runTime)
				{
					string exePath=dRow["exePath"].ToString();
					AppLauncher launcher=new AppLauncher(exePath);
					new Thread(new ThreadStart(launcher.runApp)).Start();
					// Update the next run time
					string strInterval=dRow["repeat"].ToString().ToUpper();
					switch(strInterval)
					{
						case "D":
							runTime=runTime.AddDays(1);
							break;
						case "W":
							runTime=runTime.AddDays(7);
							break;
						case "M":
							runTime=runTime.AddMonths(1);
							break;
					}
					dRow["time"]=runTime.ToString(formatString);
					dsTasks.AcceptChanges();
					StreamWriter sWrite=new StreamWriter(configPath);
					XmlTextWriter xWrite=new XmlTextWriter(sWrite);
					dsTasks.WriteXml(xWrite, XmlWriteMode.WriteSchema);
					xWrite.Close();
				}
			}
		}

		public AppScheduler()
		{
			// This call is required by the Windows.Forms Component Designer.
			InitializeComponent();

			// TODO: Add any initialization after the InitComponent call
		}

		// The main entry point for the process
		static void Main()
		{
			System.ServiceProcess.ServiceBase[] ServicesToRun;
	
			// More than one user Service may run within the same process. To add
			// another service to this process, change the following line to
			// create a second service object. For example,
			//
			//   ServicesToRun = new System.ServiceProcess.ServiceBase[] {new Service1(), new MySecondUserService()};
			//
			ServicesToRun = new System.ServiceProcess.ServiceBase[] { new AppScheduler() };

			System.ServiceProcess.ServiceBase.Run(ServicesToRun);
		}

		/// <summary> 
		/// Required method for Designer support - do not modify 
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
			// 
			// AppScheduler
			// 
			this.CanPauseAndContinue = true;
			this.ServiceName = "Application Scheduler";

		}

		/// <summary>
		/// Clean up any resources being used.
		/// </summary>
		protected override void Dispose( bool disposing )
		{
			if( disposing )
			{
				if (components != null) 
				{
					components.Dispose();
				}
			}
			base.Dispose( disposing );
		}

		/// <summary>
		/// Set things in motion so your service can do its work.
		/// </summary>
		protected override void OnStart(string[] args)
		{
			// TODO: Add code here to start your service.
			configPath=ConfigurationSettings.AppSettings["configpath"];
			try
			{
				XmlTextReader xRead=new XmlTextReader(configPath);
				XmlValidatingReader xvRead=new XmlValidatingReader(xRead);
				xvRead.ValidationType=ValidationType.DTD;
				dsTasks.ReadXml(xvRead);
				xvRead.Close();
				xRead.Close();
			}
			catch(Exception)
			{
				ServiceController srvcController=new ServiceController(ServiceName);
				srvcController.Stop();
			}
			_timer.Interval=30000;
			_timer.Elapsed+=new ElapsedEventHandler(timeElapsed);
			_timer.Start();
		}
 
		/// <summary>
		/// Stop this service.
		/// </summary>
		protected override void OnStop()
		{
			// TODO: Add code here to perform any tear-down necessary to stop your service.
		}
	}
}
