using System;

namespace RemedyRoom.FceAutomation.Services.DomainObjects.WorkSamples
{
    public class ThurstonWorkSample : WorkSampleBase
    {
        public override bool IsAccuracyJobMatch()
        {
            return false;
        }

        public override bool IsSpeedJobMatch()
        {
            return false;
        }
    }
}