using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;

namespace StockCSV.Mechanism
{
    public class JobManager
    {

        public static bool IsRealClass(Type testType)
        {
            return testType.IsAbstract == false
                   && testType.IsGenericTypeDefinition == false
                   && testType.IsInterface == false;
        }

        public void ExecuteAllJobs()
        {
            Console.WriteLine("Begin Method");

            try
            {
                var jobs = GetAllTypesImplementingInterface(typeof(Job));
                if (jobs != null && jobs.Any())
                {
                    foreach (var job in jobs)
                    {
                        if (IsRealClass(job))
                        {
                            try
                            {
                                var instanceJob = (Job)Activator.CreateInstance(job);
                                Console.WriteLine($"The Job \"{instanceJob.GetEndpoint()}\" has been instantiated successfully.");
                                var thread = new Thread(instanceJob.ExecuteJob);
                                thread.Start();
                                Console.WriteLine($"The Job \"{instanceJob.GetEndpoint()}\" has its thread started successfully.");
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine($"The Job \"{job.Name}\" could not be instantiated or executed.", ex);
                            }
                        }
                        else
                        {
                            Console.WriteLine($"The Job \"{job.FullName}\" cannot be instantiated.");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error has occured while instantiating or executing Jobs for the Scheduler Framework.", ex);
            }

            Console.WriteLine("End Method");
        }

        private static IEnumerable<Type> GetAllTypesImplementingInterface(Type desiredType)
        {
            return AppDomain.CurrentDomain.GetAssemblies().SelectMany(assembly => assembly.GetTypes())
                .Where(desiredType.IsAssignableFrom);
        }
    }
}
