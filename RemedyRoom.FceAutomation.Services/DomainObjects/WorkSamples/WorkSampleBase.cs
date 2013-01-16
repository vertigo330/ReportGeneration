using System;

namespace RemedyRoom.FceAutomation.Services.DomainObjects.WorkSamples
{
    public abstract class WorkSampleBase
    {
        public string Name { get; set; }
        public TimeSpan Speed { get; set; }
        public int Accuracy { get; set; }
        
        public virtual bool IsAccuracyJobMatch()
        {
            throw new NotImplementedException();
        }

        public virtual bool IsSpeedJobMatch()
        {
            throw new NotImplementedException();
        }
    }
}