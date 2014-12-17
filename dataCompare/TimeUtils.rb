
module TimeUtils

  @@date = nil
  @@month = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"]

  # Returns the module's internal date that, if previously set, will be used for module calculations.
  #
  # added 0.1.0
  def get_date
    @@date = system_date
  end

  # Set the module's internal date that will be used for module calculations.
  #
  # added 0.1.0
  def set_current_date(current_date)
    validate_date(current_date)

    @@date = current_date
  end

  # Removes the module's internal date.
  #
  # added 0.1.1
  def reset_current_date
    @@date = nil
  end

  # Returns the current date of the system.

  def system_date
    format_date(Time.now)
  end



  # Returns the cardinal position in the year of a given month in mm format.

  def month_cardinal(month)
    validate_month(month)

    digit = @@months.index(month) + 1
    if (digit < 10)
      digit = "0#{digit}"
    end

    "#{digit}"
  end



  private

  # Returns the Time object that will be used for calculations within the module. The date used to initialize the
  # Time is either the date set by set_current_date or the date as determined by the current system time.
  #
  # added 0.1.0
  def retrieve_time
    if @@date
      format_time(@@date)
    else
      Time.now
    end
  end

  # Returns a Time object initialized to the given mm/dd//yyyy date
  #
  # added 0.1.0
  def format_time(date)
    year = date[6, 4]
    month = date[0, 2]
    day = date[3, 2]

    Time.new(year, month, day)
  end

  # Returns a mm/dd/yyyy formatted String based on the given Time argument.
  #
  # added 0.1.0
  def format_date(time)
    time.strftime("%m/%d/%Y").gsub("/", "-")
  end


  # Validates that a month is a real month.
  #
  # added 0.1.2
  def validate_month(month)
    raise(ArgumentError, "Month '#{month}' is not in a valid month.") unless @@months.include?(month)
  end


end
