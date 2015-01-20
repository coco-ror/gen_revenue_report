# encoding: UTF-8
module Axlsx

  # The LineChart is a two dimentional line chart (who would have guessed?) that you can add to your worksheet.
  # @example Creating a chart
  #   # This example creates a line in a single sheet.
  #   require "rubygems" # if that is your preferred way to manage gems!
  #   require "axlsx"
  #
  #   p = Axlsx::Package.new
  #   ws = p.workbook.add_worksheet
  #   ws.add_row ["This is a chart with no data in the sheet"]
  #
  #   chart = ws.add_chart(Axlsx::LineChart, :start_at=> [0,1], :end_at=>[0,6], :title=>"Most Popular Pets")
  #   chart.add_series :data => [1, 9, 10], :labels => ["Slimy Reptiles", "Fuzzy Bunnies", "Rottweiler"]
  #
  # @see Worksheet#add_chart
  # @see Worksheet#add_row
  # @see Chart#add_series
  # @see Series
  # @see Package#serialize
  class CombineChart < Chart

    # the category axis
    # @return [CatAxis]
    def cat_axis
      axes[:cat_axis]
    end
    alias :catAxis :cat_axis

    # the category axis
    # @return [ValAxis]
    def val_axis
      axes[:val_axis]
    end
    alias :valAxis :val_axis

    # the secondary category axis
    # @return [sec_cat_axis]
    def sec_cat_axis
      axes[:sec_cat_axis]
    end
    alias :secCatAxis :sec_cat_axis

    # the secondary values axis
    # @return [sec_val_axis]
    def sec_val_axis
      axes[:sec_val_axis]
    end
    alias :secValAxis :sec_val_axis

    # The direction of the bars in the chart
    # must be one of [:bar, :col]
    # @return [Symbol]
    def bar_dir
      @bar_dir ||= :bar
    end
    alias :barDir :bar_dir


    # space between bar or column clusters, as a percentage of the bar or column width.
    # @return [String]
    attr_reader :gap_depth
    alias :gapDepth :gap_depth

    # space between bar or column clusters, as a percentage of the bar or column width.
    # @return [String]
    def gap_width
      @gap_width ||= 150
    end
    alias :gapWidth :gap_width

    #grouping for a column, line, or area chart.
    # must be one of  [:percentStacked, :clustered, :standard, :stacked]
    # @return [Symbol]
    def grouping
      @grouping ||= :clustered
    end

    # The shabe of the bars or columns
    # must be one of  [:cone, :coneToMax, :box, :cylinder, :pyramid, :pyramidToMax]
    # @return [Symbol]
    def shape
      @shape ||= :box
    end

    # validation regex for gap amount percent
    GAP_AMOUNT_PERCENT = /0*(([0-9])|([1-9][0-9])|([1-4][0-9][0-9])|500)%/

    # must be one of  [:percentStacked, :clustered, :standard, :stacked]

    # @return [Symbol]
    attr_reader :grouping

    # Creates a new line chart object
    # @param [GraphicFrame] frame The workbook that owns this chart.
    # @option options [Cell, String] title
    # @option options [Boolean] show_legend
    # @option options [Symbol] grouping
    # @see Chart
    def initialize(frame, options={})
      @vary_colors = true
      @grouping = :standard
      @gap_width, @gap_depth, @shape = nil, nil, nil
      super(frame, options)
      @series_type = LineSeries
      @d_lbls = nil
    end

    def add_series(type, options={})
      if type ==  'line'
        LineSeries.new(self, options)
      else
        BarSeries.new(self, options)
      end
      @series.last
    end

    # The direction of the bars in the chart
    # must be one of [:bar, :col]
    def bar_dir=(v)
      RestrictionValidator.validate "BarChart.bar_dir", [:bar, :col], v
      @bar_dir = v
    end
    alias :barDir= :bar_dir=

    #grouping for a column, line, or area chart.
    # must be one of  [:percentStacked, :clustered, :standard, :stacked]
    # @see grouping
    def grouping=(v)
      RestrictionValidator.validate "BarChart.grouping", [:percentStacked, :standard, :stacked], v
      @grouping = v
    end

    # space between bar or column clusters, as a percentage of the bar or column width.
    def gap_width=(v)
      RegexValidator.validate "BarChart.gap_width", GAP_AMOUNT_PERCENT, v
      @gap_width=(v)
    end
    alias :gapWidth= :gap_width=

    # space between bar or column clusters, as a percentage of the bar or column width.
    def gap_depth=(v)
      RegexValidator.validate "BarChart.gap_didth", GAP_AMOUNT_PERCENT, v
      @gap_depth=(v)
    end
    alias :gapDepth= :gap_depth=

    # The shabe of the bars or columns
    # must be one of  [:cone, :coneToMax, :box, :cylinder, :pyramid, :pyramidToMax]
    def shape=(v)
      RestrictionValidator.validate "BarChart.shape", [:cone, :coneToMax, :box, :cylinder, :pyramid, :pyramidToMax], v
      @shape = v
    end


    # Serializes the object
    # @param [String] str
    # @return [String]
    def to_xml_string(str = '')
      if @series.all? {|s| s.on_primary_axis} then
        # Only a primary val axis
        super(str) do
          str << ('<c:lineChart>')
          str << ('<c:grouping val="' << grouping.to_s << '"/>')
          str << ('<c:varyColors val="' << vary_colors.to_s << '"/>')
          @series.each { |ser| ser.to_xml_string(str) }
          @d_lbls.to_xml_string(str) if @d_lbls
          yield if block_given?
          axes.to_xml_string(str, :ids => true)
          str << ('</c:lineChart>')
          axes.to_xml_string(str)
        end
      else
        # Two value axes
        super(str) do
          # First axis
          str << ('<c:barChart>')
          str << ('<c:barDir val="' << bar_dir.to_s << '"/>')
          str << ('<c:grouping val="' << grouping.to_s << '"/>')
          str << ('<c:varyColors val="' << vary_colors.to_s << '"/>')
          @series.select {|s| !s.on_primary_axis}.each { |s| s.to_xml_string(str) }
          @d_lbls.to_xml_string(str) if @d_lbls
          yield if block_given?
          str << ('<c:gapWidth val="' << @gap_width.to_s << '"/>') unless @gap_width.nil?
          str << ('<c:gapDepth val="' << @gap_depth.to_s << '"/>') unless @gap_depth.nil?
          str << ('<c:shape val="' << @shape.to_s << '"/>') unless @shape.nil?
          str << ('<c:axId val="' << axes[:cat_axis].id.to_s << '"/>')
          str << ('<c:axId val="' << axes[:val_axis].id.to_s << '"/>')
          axes.to_xml_string(str, :ids => true)
          str << ('</c:barChart>')

          # Secondary axis
          str << ('<c:lineChart>')
          str << ('<c:grouping val="' << grouping.to_s << '"/>')
          str << ('<c:varyColors val="' << vary_colors.to_s << '"/>')
          @series.select {|s| s.on_primary_axis}.each { |s| s.to_xml_string(str) }
          @d_lbls.to_xml_string(str) if @d_lbls
          yield if block_given?
          str << ('<c:axId val="' << axes[:sec_cat_axis].id.to_s << '"/>')
          str << ('<c:axId val="' << axes[:sec_val_axis].id.to_s << '"/>')
          str << ('</c:lineChart>')

          # The axes
          axes.to_xml_string(str)
        end
      end
    end

    # The axes for this chart. LineCharts have a category and value
    # axis.
    # @return [Axes]
    def axes
      if @axes.nil? then
        # add the normal axes
        @axes = Axes.new(:cat_axis => CatAxis, :val_axis => ValAxis)

        # add the secondary axes if needed
        if @series.any? {|s| !s.on_primary_axis} then
          if @axes[:sec_cat_axis].nil? then
            @axes.add_axis(:sec_cat_axis, Axlsx::CatAxis)
            sec_cat_axis = @axes[:sec_cat_axis]
            sec_cat_axis.ax_pos = :b
            sec_cat_axis.delete = 1
            sec_cat_axis.gridlines = false
          end
          if @axes[:sec_val_axis].nil? then
            @axes.add_axis(:sec_val_axis, Axlsx::ValAxis)
            sec_val_axis = @axes[:sec_val_axis]
            sec_val_axis.ax_pos = :r
            sec_val_axis.gridlines = false
            sec_val_axis.crosses = :max
            sec_val_axis.format_code = '0.00%'
          end
        end
      end

      # return
      @axes
    end
  end
end