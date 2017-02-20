using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Quartz;
using Quartz.Impl;

namespace ExcelPptTest
{
    public partial class QuartzTest : Form , IJob
    {
        IScheduler sched = null;

        public QuartzTest()
        {
            InitializeComponent();
        }

        public void Execute(IJobExecutionContext context)
        {
            MessageBox.Show("job start");
        }

        private void QuartzTest_Load(object sender, EventArgs e)
        {
            // construct a scheduler factory
            ISchedulerFactory schedFact = new StdSchedulerFactory();

            // get a scheduler
            sched = schedFact.GetScheduler();
            sched.Start();

            IJobDetail job = JobBuilder.Create<QuartzTest>()
                .WithIdentity("myJob", "group1")
                .Build();

            ITrigger trigger = TriggerBuilder.Create()
               .WithDailyTimeIntervalSchedule
                 (s =>
                    s.WithIntervalInHours(24)
                   .OnEveryDay()
                   .StartingDailyAt(TimeOfDay.HourAndMinuteOfDay(19, 07))
                 )
               .Build();

            sched.ScheduleJob(job, trigger);
        }

        private void QuartzTest_FormClosing(object sender, FormClosingEventArgs e)
        {
            sched.Shutdown();
        }
    }

    //public class LoggingJob : IJob
    //{
    //    public void Execute(IJobExecutionContext context)
    //    {
    //        MessageBox.Show("job start");
    //    }
    //}
}
