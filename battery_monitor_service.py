import os
import wmi
import time
import win32serviceutil
import win32service
import win32event
import servicemanager

class BatteryMonitorService(win32serviceutil.ServiceFramework):
    _svc_name_ = "BatteryMonitorService"
    _svc_display_name_ = "Battery Monitor Service"
    _svc_description_ = "A service that monitors the laptop's battery and puts it to sleep if it is not charging."

    def __init__(self, args):
        win32serviceutil.ServiceFramework.__init__(self, args)
        self.hWaitStop = win32event.CreateEvent(None, 0, 0, None)
        self.running = True

    def SvcStop(self):
        # Report that the service is stopping
        self.ReportServiceStatus(win32service.SERVICE_STOP_PENDING)
        win32event.SetEvent(self.hWaitStop)
        self.running = False

    def SvcDoRun(self):
        # Report that the service is running
        self.ReportServiceStatus(win32service.SERVICE_RUNNING)
        servicemanager.LogInfoMsg("Battery Monitor Service is running.")
        
        # Call the method to monitor the battery
        self.monitor_battery()

    def check_battery_status(self):
        # Check if the battery is charging
        c = wmi.WMI()
        for battery in c.Win32_Battery():
            return battery.BatteryStatus == 2  # 2 = Charging

    def put_laptop_to_sleep(self):
        # Put the system to sleep if not charging
        os.system("rundll32.exe powrprof.dll,SetSuspendState 0,1,0")

    def monitor_battery(self):
        while self.running:
            is_charging = self.check_battery_status()
            if not is_charging:
                servicemanager.LogInfoMsg("Laptop is not charging. Going to sleep...")
                self.put_laptop_to_sleep()
                break
            time.sleep(2)

if __name__ == "__main__":
    win32serviceutil.HandleCommandLine(BatteryMonitorService)
