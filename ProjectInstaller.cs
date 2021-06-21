using System;
using System.Collections;
using System.ComponentModel;
using System.Configuration.Install;

namespace AppScheduler
{
	/// <summary>
	/// Summary description for ProjectInstaller.
	/// </summary>
	[RunInstaller(true)]
	public class ProjectInstaller : System.Configuration.Install.Installer
	{
		private System.ServiceProcess.ServiceProcessInstaller spInstaller;
		private System.ServiceProcess.ServiceInstaller srvcInstaller;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public ProjectInstaller()
		{
			// This call is required by the Designer.
			InitializeComponent();

			// TODO: Add any initialization after the InitializeComponent call
		}

		/// <summary> 
		/// Clean up any resources being used.
		/// </summary>
		protected override void Dispose( bool disposing )
		{
			if( disposing )
			{
				if(components != null)
				{
					components.Dispose();
				}
			}
			base.Dispose( disposing );
		}


		#region Component Designer generated code
		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
			this.spInstaller = new System.ServiceProcess.ServiceProcessInstaller();
			this.srvcInstaller = new System.ServiceProcess.ServiceInstaller();
			// 
			// spInstaller
			// 
			this.spInstaller.Account = System.ServiceProcess.ServiceAccount.LocalSystem;
			this.spInstaller.Password = null;
			this.spInstaller.Username = null;
			// 
			// srvcInstaller
			// 
			this.srvcInstaller.ServiceName = "AppScheduler";
			this.srvcInstaller.StartType = System.ServiceProcess.ServiceStartMode.Automatic;
			// 
			// ProjectInstaller
			// 
			this.Installers.AddRange(new System.Configuration.Install.Installer[] {
																					  this.spInstaller,
																					  this.srvcInstaller});

		}
		#endregion
	}
}
