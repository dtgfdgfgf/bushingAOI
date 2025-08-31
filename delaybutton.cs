using System.Threading;
using System.Threading.Tasks;
using PLC;

namespace System.Windows.Forms
{
    public class DelayButton : Button
    {
        public DelayButton() { }
        public int Interval { get; set; } = 15000;
        public bool UseEnable { get; set; } = true;
        public bool before { get; set; } = true;
        public bool after { get; set; } = true;
        protected override void OnClick(EventArgs e)
        {
            if (!app.offline)
            {
                var Handle = new ManualResetEvent(false);
                PLC_ModBus.SetPoint(1, BoolUnit.M, 0, false, false, Handle);
                Handle.WaitOne();
            }

            if (UseEnable)
            {
                this.Enabled = false;
            }

            if (before)
            {
                Task.Run(async () =>
                {
                    await Task.Delay(this.Interval);

                    this.Invoke(new Action(() =>
                    {
                        base.OnClick(e);

                        //if (UseEnable)
                        //{
                        //    this.Enabled = true;
                        //}
                    }));
                });

            }
            else if (after)
            {                
                Task.Run(async () =>
                {
                    this.Invoke(new Action(() =>
                    {
                        base.OnClick(e);                        
                    }));                    
                    await Task.Delay(this.Interval);
                    this.Invoke(new Action(() =>
                    {
                        base.Enabled = true ;
                    }));
                });                
                
            }
        }
    }
}
