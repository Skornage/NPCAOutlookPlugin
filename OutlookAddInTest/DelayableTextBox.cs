using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Threading.Tasks;

namespace OutlookAddInTest
{
	public class DelayableTextBox : TextBox
	{
		private Timer m_delayedTextChangedTimer;
		private event EventHandler delayedTextChangedTimerTickHandler;

		public event EventHandler DelayedTextChanged;

		public DelayableTextBox()
			: base()
		{
			this.DelayedTextChangedTimeout = 1000; // 1.0 seconds
		}

		protected override void Dispose(bool disposing)
		{
			if (m_delayedTextChangedTimer != null)
			{
				m_delayedTextChangedTimer.Stop();
				if (disposing)
					m_delayedTextChangedTimer.Dispose();
			}

			base.Dispose(disposing);
		}

		public int DelayedTextChangedTimeout { get; set; }

		public virtual List<Result> OnDelayedTextChanged(EventArgs e)
		{
			List<Result> results = new List<Result>();
			if (this.DelayedTextChanged != null)
				this.DelayedTextChanged(this, e);
			if (this.Text != null)
			{
				System.Diagnostics.Trace.WriteLine("Searching");
				String[] query = parseInputText(this.Text);
				if (query != null)
				{
					results = JsonGetter.GetSearchResults(query);
				}
			}
			return results;
		}

		private String[] parseInputText(String inputText) 
		{
			if (inputText == "")
			{
				return null;
			}
			inputText = inputText.TrimEnd(new char[] {' ', '\n', '\r'});
			inputText = inputText.TrimStart(new char[] { ' ', '\n', '\r' });
			if (inputText != null)
			{
				String[] results = new String[0];
				System.Diagnostics.Trace.WriteLine("Text:"+inputText+"B");
				results = inputText.Split(' ');
				return results;
			}
			return null;
		}

		protected override void OnTextChanged(EventArgs e)
		{
			this.InitializeDelayedTextChangedEvent();
			base.OnTextChanged(e);
		}

		public void InitializeDelayedTextChangedEvent()
		{
			if (m_delayedTextChangedTimer != null)
				m_delayedTextChangedTimer.Stop();

			if (m_delayedTextChangedTimer == null || m_delayedTextChangedTimer.Interval != this.DelayedTextChangedTimeout)
			{
				m_delayedTextChangedTimer = new Timer();
				m_delayedTextChangedTimer.Tick += delayedTextChangedTimerTickHandler;
				m_delayedTextChangedTimer.Interval = this.DelayedTextChangedTimeout;
			}

			m_delayedTextChangedTimer.Start();
		}

		private void HandleDelayedTextChangedTimerTick(object sender, EventArgs e)
		{
			Timer timer = sender as Timer;
			timer.Stop();

			this.OnDelayedTextChanged(EventArgs.Empty);
		}

		public void setDelayedTextChangedTimerTickHandler(EventHandler handler) {
			this.delayedTextChangedTimerTickHandler = handler;
		}
	}
}
