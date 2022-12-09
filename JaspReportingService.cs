using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Diagnostics;
using System.ServiceProcess;
using System.Timers;

namespace jaspReportingService
{
	public partial class JaspReportingService : ServiceBase
	{
		private Process _jasp;
		private Timer	_timer;
		public JaspReportingService()
		{
			//
			InitializeComponent();
		}

		protected override void OnStart(string[] args)
		{
			WriteEventLog(false, "Service started");
			string msg;
			try
			{
				System.Threading.Thread.Sleep(3000);
				WriteEventLog(StartJasp(out msg), msg);
				
				_timer = new Timer();
				_timer.Interval = 1000 * 10; //For development 10 sec is already on the slow end, while for release fast?
				_timer.Elapsed	+= OnTimedEvent;
				_timer.AutoReset = true;
				_timer.Start();
			}
			catch(Exception e)
			{
				WriteEventLog(true, e.Message.ToString());
			}
		}

		private void OnTimedEvent(Object source, System.Timers.ElapsedEventArgs e)
		{
			if(_jasp != null && _jasp.Responding && _jasp.StandardError.EndOfStream)
				return;

			if(_jasp == null || _jasp.HasExited)
			{
				_timer.Stop();
				WriteExitEvent();
				Stop();
				return;
			}

			WriteEventLog(false, "JASP had stderr output: " + _jasp == null ? "JASP never started!" : _jasp.StandardError.ReadToEnd());
		}

		protected override void OnStop()
		{
			WriteEventLog(false, "Service stopped");

			if(_jasp != null)
			{
				if(_jasp.Responding)
					_jasp.Close();

				if(!_jasp.HasExited)
					_jasp.Kill();
			}

			WriteExitEvent();
		}

		private void WriteExitEvent()
		{
			string exitMsg = _jasp == null ? "JASP never started..." : "JASP exited with exitcode " + _jasp.ExitCode.ToString() + " and stderror: " + _jasp.StandardError.ReadToEnd();
			WriteEventLog(_jasp == null ? true : _jasp.ExitCode != 0, exitMsg);
		}

		private void WriteEventLog(bool error, string msg)
		{
			const int maxEntryLength = 32000 ; // The msgs really should be shorter than this, but *just* in case:
			int splitMsg = (int)Math.Ceiling(msg.Length / (float)maxEntryLength);

			for(int split = 0; split<splitMsg; split++)
			{
				int curPos = maxEntryLength * split;
				EventLog.WriteEntry(
					(splitMsg > 1 ? ("Part #" + (1+split).ToString() + " of " + splitMsg + ": ") : "") 
					+ msg.Substring(startIndex: curPos, length: Math.Min(msg.Length - curPos, maxEntryLength))
					, !error ? EventLogEntryType.Information : EventLogEntryType.Error);
			}
		}

		private bool StartJasp(out string msg)
		{
			msg = "???";

			System.Collections.Specialized.NameValueCollection app = ConfigurationManager.AppSettings;
			
			Process cmd = new Process();
			cmd.StartInfo.WorkingDirectory			= app.Get("JaspDesktopLocation");
			cmd.StartInfo.FileName					= "JASP.exe";
			cmd.StartInfo.RedirectStandardInput		= false;
			cmd.StartInfo.RedirectStandardOutput	= true;
			cmd.StartInfo.CreateNoWindow			= true;
			cmd.StartInfo.UseShellExecute			= false;
			cmd.StartInfo.Arguments					= "--report " + app.Get("ReportOutputFolder") + " " + app.Get("ReportJaspFile");
			
			bool itWorked = false;
			try 
			{
				itWorked = cmd.Start();
				msg = "Jasp is running.";
				_jasp = cmd;
			}
			catch(System.InvalidOperationException e) 
			{ 
				msg =	"No file name was specified in the System.Diagnostics.Process component's System.Diagnostics.Process.StartInfo.-or-\n" +
						"The System.Diagnostics.ProcessStartInfo.UseShellExecute member of the System.Diagnostics.Process.StartInfo\n" +
						"property is true while System.Diagnostics.ProcessStartInfo.RedirectStandardInput,\n" +
						"System.Diagnostics.ProcessStartInfo.RedirectStandardOutput, or System.Diagnostics.ProcessStartInfo.RedirectStandardError\n" +
						"is true.\nError was: " + e.Message.ToString();
		
			}
			catch(System.ComponentModel.Win32Exception e) 
			{ 
				msg = "There was an error in opening the associated file: " + e.Message.ToString(); 
			}

			return itWorked;
		}
	}
}
