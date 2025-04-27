use std::io::Write;
use std::io::{self};
use std::time::Duration;
use std::time::SystemTime;
use std::time::UNIX_EPOCH;
use windows::Devices::Power::Battery;
use windows::Devices::Power::BatteryReport;
use windows::Foundation::DateTime;
use windows::Foundation::TimeSpan;
use windows::Foundation::TypedEventHandler;
use windows::Globalization::DateTimeFormatting::DateTimeFormatter;
use windows::Win32::System::Com::COINIT_MULTITHREADED;
use windows::Win32::System::Com::CoInitializeEx;
use windows::Win32::System::Com::CoUninitialize;

fn format_date_time(date_time: &DateTime) -> String {
    let formatter = DateTimeFormatter::LongDate().unwrap();
    let owned_date_time = DateTime {
        UniversalTime: date_time.UniversalTime,
    };
    formatter.Format(owned_date_time).unwrap().to_string()
}

fn system_time_to_datetime(time: SystemTime) -> DateTime {
    let duration = time.duration_since(UNIX_EPOCH).unwrap();
    let unix_time = duration.as_secs() as i64;
    // Convert Unix timestamp to Windows FILETIME (100-nanosecond intervals since January 1, 1601)
    let windows_ticks = (unix_time + 11644473600) * 10000000;
    DateTime {
        UniversalTime: windows_ticks,
    }
}

fn format_severity(severity: f64) -> String {
    format!("{:.2}", severity)
}

fn show_forecast() {
    let battery = Battery::AggregateBattery().unwrap();
    let report = battery.GetReport().unwrap();

    if let Ok(remaining_capacity) = report.RemainingCapacityInMilliwattHours() {
        let remaining_capacity = remaining_capacity.GetInt32().unwrap();
        let full_capacity = report
            .FullChargeCapacityInMilliwattHours()
            .map(|c| c.GetInt32().unwrap_or(100))
            .unwrap_or(100);
        let charge_rate = report
            .ChargeRateInMilliwatts()
            .map(|r| r.GetInt32().unwrap_or(0))
            .unwrap_or(0);
        let current_time = system_time_to_datetime(SystemTime::now());

        println!(
            "Current battery status at {}",
            format_date_time(&current_time)
        );
        println!(
            "Remaining capacity: {}%",
            (remaining_capacity as f64 / full_capacity as f64 * 100.0) as i32
        );
        println!("Charge rate: {} mW", charge_rate);
        println!();

        if charge_rate > 0 {
            let time_to_full = (full_capacity - remaining_capacity) as f64 / charge_rate as f64;
            let duration_ticks = (time_to_full * 3600.0 * 10000000.0) as i64;
            let full_time = DateTime {
                UniversalTime: current_time.UniversalTime + duration_ticks,
            };
            println!(
                "Estimated time until full: {}",
                format_date_time(&full_time)
            );
        } else if charge_rate < 0 {
            let time_to_empty = remaining_capacity as f64 / (-charge_rate as f64);
            let duration_ticks = (time_to_empty * 3600.0 * 10000000.0) as i64;
            let empty_time = DateTime {
                UniversalTime: current_time.UniversalTime + duration_ticks,
            };
            println!(
                "Estimated time until empty: {}",
                format_date_time(&empty_time)
            );
        }
    } else {
        println!("Could not get battery capacity information");
    }
    println!();
}

fn get_battery_level(battery: &Battery) -> f64 {
    let report = battery.GetReport().unwrap();
    if let Ok(remaining) = report.RemainingCapacityInMilliwattHours() {
        let remaining = remaining.GetInt32().unwrap_or(0);
        let full = report
            .FullChargeCapacityInMilliwattHours()
            .map(|c| c.GetInt32().unwrap_or(100))
            .unwrap_or(100);
        remaining as f64 / full as f64
    } else {
        0.0
    }
}

fn find_best(look_ahead: Duration, require_charging: bool) {
    let battery = Battery::AggregateBattery().unwrap();
    let current_level = get_battery_level(&battery);

    let report = battery.GetReport().unwrap();
    let charge_rate = report
        .ChargeRateInMilliwatts()
        .map(|r| r.GetInt32().unwrap_or(0))
        .unwrap_or(0);

    if require_charging && charge_rate <= 0 {
        println!("Battery is not charging. Cannot find optimal charging time.");
        return;
    }

    let start_time = system_time_to_datetime(SystemTime::now());
    let duration_ticks = look_ahead.as_secs() as i64 * 10000000;
    let end_time = DateTime {
        UniversalTime: start_time.UniversalTime + duration_ticks,
    };

    if charge_rate > 0 {
        let full_capacity = report
            .FullChargeCapacityInMilliwattHours()
            .map(|c| c.GetInt32().unwrap_or(100))
            .unwrap_or(100);
        let time_to_full = ((1.0 - current_level) * full_capacity as f64) / charge_rate as f64;
        let full_time = DateTime {
            UniversalTime: start_time.UniversalTime + (time_to_full * 3600.0 * 10000000.0) as i64,
        };

        if full_time.UniversalTime <= end_time.UniversalTime {
            println!(
                "Battery will be fully charged at {} (in {:.1} hours)",
                format_date_time(&full_time),
                time_to_full
            );
        } else {
            println!(
                "Battery will reach {:.1}% at {} (in {} hours)",
                current_level * 100.0 + (charge_rate as f64 * look_ahead.as_secs_f64() / 3600.0),
                format_date_time(&end_time),
                look_ahead.as_secs_f64() / 3600.0
            );
        }
    } else {
        println!(
            "Battery is not charging. Current level: {:.1}%",
            current_level * 100.0
        );
    }
    println!();
}

fn perform_forecast_calculations() {
    show_forecast();

    println!("Battery charging forecast for the next 10 hours (requiring charging):");
    find_best(Duration::from_secs(10 * 3600), true);

    println!("Battery status forecast for the next 10 hours:");
    find_best(Duration::from_secs(10 * 3600), false);
}

fn main() {
    unsafe {
        CoInitializeEx(None, COINIT_MULTITHREADED).unwrap();
    }

    perform_forecast_calculations();

    let battery = Battery::AggregateBattery().unwrap();
    let handler = TypedEventHandler::new(|_, _| {
        perform_forecast_calculations();
        Ok(())
    });
    let _token = battery.ReportUpdated(&handler).unwrap();

    println!("Waiting for battery status changes...");
    println!("Press Enter to exit the program.");

    let mut input = String::new();
    io::stdin().read_line(&mut input).unwrap();

    battery.RemoveReportUpdated(_token).unwrap();

    unsafe {
        CoUninitialize();
    }
}
