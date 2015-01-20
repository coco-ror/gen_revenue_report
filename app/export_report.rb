$LOAD_PATH.unshift "#{File.dirname(__FILE__)}/../lib"
require 'axlsx'
require 'json'
require './constants'

report_file = File.read('assets/data/report.json')
@report_data = eval(report_file)

p = Axlsx::Package.new
p.use_autowidth = false
wb = p.workbook

#define your regular styles
styles = wb.styles

st_header = styles.add_style :sz => 26, :b => true, :u => true, :border => Axlsx::STYLE_THIN_BORDER
st_sub_header = styles.add_style :sz => 16, :fg_color => Constants::PINK_COLOR
st_footer = styles.add_style :sz => 12, :bg_color => Constants::BLUE_COLOR, :fg_color => Constants::WHITE_COLOR, :alignment => {:vertical => :center},:border => Axlsx::STYLE_THIN_BORDER
st_date = styles.add_style :format_code => "yyyy年mm月", :sz => 24, :alignment => {:horizontal => :center, :vertical => :center}
st_number = styles.add_style :format_code => '#,###,##0', :sz => 24, :alignment => {:horizontal => :center, :vertical => :center}
st_currency = styles.add_style :format_code => '¥#,###,##0', :sz => 24, :alignment => {:horizontal => :center, :vertical => :center}
st_percent = styles.add_style :format_code => '0.00%', :sz => 24, :alignment => {:horizontal => :center, :vertical => :center}
st_title = styles.add_style :sz => 20
st_tb_head = styles.add_style :sz => 16, :bg_color => Constants::BLUE_COLOR, :fg_color => Constants::WHITE_COLOR, :alignment => {:horizontal => :center, :vertical => :center},:border => Axlsx::STYLE_THIN_BORDER
st_tb_body = styles.add_style :sz => 24, :alignment => {:horizontal => :center, :vertical => :center}
st_summary = styles.add_style :sz => 16, :b => true, :bg_color => Constants::SMOOTH_BLUE_COLOR, :alignment => {:horizontal => :left, :vertical => :center},:border => Axlsx::STYLE_THIN_BORDER

st_amount_currency = styles.add_style :format_code => '¥#,###,##0', :sz => 24, :bg_color => Constants::YELLOW_COLOR, :alignment => {:horizontal => :center, :vertical => :center}
st_amount_percent = styles.add_style :format_code => '0.00%', :sz => 24, :bg_color => Constants::YELLOW_COLOR, :alignment => {:horizontal => :center, :vertical => :center}
st_bg_clear = styles.add_style(:bg_color => Constants::WHITE_COLOR)

st_profitable = styles.add_style(:bg_color => Constants::YELLOW_COLOR, :type => :dxf)


WB_HEAD_TITLE = @report_data[:head_title]
WB_FS_SUBHEAD_TITLE = @report_data[:head_ft_summary]
WB_SE_SUBHEAD_TITLE = @report_data[:head_se_summary]
WB_TH_SUBHEAD_TITLE = @report_data[:head_th_summary]
WB_FOOTER_TITLE = @report_data[:footer_title]

REPORT_AMOUNT = @report_data[:amount]
REPORT_DETAILS = @report_data[:details]
REPORT_FIR_DATA = @report_data[:first_data]

SEC_REPORT_DATA = @report_data[:second_data]
THI_REPORT_DATA = @report_data[:third_data]
FOR_REPORT_DATA = @report_data[:fourth_data]
FIF_REPORT_DATA = @report_data[:fifth_data]

#define your first sheet
wb.add_worksheet(:name => Constants::SHEET_FIR_NAME) do |ws|
  ws.add_row
  ws.add_row [nil, WB_HEAD_TITLE], :style => st_header, :width => :auto, :height => 40

  ws.add_row [nil, WB_FS_SUBHEAD_TITLE], :style => st_sub_header, :height => 30, :width => :auto
  ws.add_row [nil, WB_SE_SUBHEAD_TITLE], :style => st_sub_header, :height => 30, :width => :auto

  ws.add_row [nil, nil, nil, nil, nil, nil, nil, nil, '総予算', 'Amazon全体 売上目標', '寄与率'], :style => st_tb_head, :height => 30

  ws.add_row [nil, Constants::FSSHEET_FIR_TITLE, nil, nil, nil, nil, nil, nil, REPORT_AMOUNT[:total_budget], REPORT_AMOUNT[:amazon_sales], REPORT_AMOUNT[:percent]], :style => [nil, st_title, nil, nil, nil, nil, nil, nil, st_amount_currency, st_amount_currency, st_amount_percent], :height => 30
  ws["A5:H5"].each { |c| c.style = st_bg_clear }

  ws.add_row Constants::FSSHEET_FTB_TITLES, :style => st_tb_head, :height => 30, :widths => [1, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30]
  ws.add_row [nil, REPORT_DETAILS[:日数], REPORT_DETAILS[:imp], REPORT_DETAILS[:click], REPORT_DETAILS[:CTR], REPORT_DETAILS[:CPC], REPORT_DETAILS[:CV], REPORT_DETAILS[:CPA], REPORT_DETAILS[:広告費], REPORT_DETAILS[:総売上], REPORT_DETAILS[:ROAS]], :height => 60, :style => [nil, st_tb_body, st_number, st_number, st_percent, st_number, st_number, st_currency, st_currency, st_number, st_percent], :width => :auto

  # define chart start
  ws.add_row
  ws.add_row [nil, Constants::FSSHEET_SEC_TITLE], :style => st_title, :height => 30
  ws.add_row [nil, Constants::FSSHEET_CHART_TITLE], :style => st_tb_head, :height => 30
  ws.merge_cells "B11:K11"

  5.times{ ws.add_row [], :height => 90 }
  ws.add_row

  ws.add_row [nil, Constants::FSSHEET_THI_TITLE], :style => st_title, :height => 30
  ws.add_row Constants::FSSHEET_STB_TITLES, :style => st_tb_head, :height => 30, :widths => [1, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30]

  REPORT_FIR_DATA[:first].each do |item|
    ws.add_row [nil, item[:週], item[:imp], item[:click], item[:CTR], item[:CPC], item[:CV], item[:CPA], item[:広告費], item[:総売上], item[:ROAS]], :height => 60, :style => [nil, st_date, st_number, st_number, st_percent, st_currency, st_number, st_currency, st_currency, st_currency, st_percent], :width => :auto
  end

  ws.add_conditional_formatting("I20:I34", { :type => :dataBar, :dxfId => st_profitable, :priority => 1, :data_bar => Axlsx::DataBar.new })
  ws.add_conditional_formatting("J20:J34", { :type => :dataBar, :dxfId => st_profitable, :priority => 1, :data_bar => Axlsx::DataBar.new })
  ws.add_conditional_formatting("K20:K34", { :type => :dataBar, :dxfId => st_profitable, :priority => 1, :data_bar => Axlsx::DataBar.new })

  ws.add_chart(Axlsx::CombineChart, :title => ' ', :bar_dir => :col) do |chart|
    chart.start_at 1, 11
    chart.end_at 11, 16
    chart.add_series 'bar', :data => ws["J20:J34"], :labels => ws["B20:B34"], :title => ws["J19"], :colors => (1..15).map{Constants::BLUE_COLOR}, :on_primary_axis => false
    chart.add_series 'line', :data => ws["K20:K34"], :labels => ws["B20:B34"], :title => ws["K19"], :color => Constants::YELLOW_COLOR, :show_marker => true
    chart.catAxis.label_rotation = -45
    chart.d_lbls.d_lbl_pos = :t
    chart.d_lbls.show_val = true
    chart.valAxis.gridlines = false
    chart.catAxis.gridlines = false
    chart.val_axis.format_code = '¥#,###,##0'
  end

  ws.add_row
  ws.add_row [nil, Constants::FSSHEET_COMMENT_TITLE], :style => st_tb_head, :height => 30
  ws.merge_cells "B36:K36"

  ws.add_row [nil, Constants::FSSHEET_LIST]
  ws.merge_cells "B37:K42"
  7.times{ ws.add_row }
  # define chart end

  ws.add_row [nil, Constants::FSSHEET_FIR_SUMMARY], :style => st_summary, :height => 30
  ws.merge_cells "B45:K45"

  # define chart start
  ws.add_row
  ws.add_row [nil, Constants::FSSHEET_FOR_TITLE], :style => st_title, :height => 30
  ws.add_row [nil, Constants::FSSHEET_CHART_TITLE], :style => st_tb_head, :height => 30
  ws.merge_cells "B48:K48"

  5.times{ ws.add_row [], :height => 90 }
  ws.add_row

  ws.add_row [nil, Constants::FSSHEET_FIF_TITLE], :style => st_title, :height => 30
  ws.add_row Constants::FSSHEET_STB_TITLES, :style => st_tb_head, :height => 30, :widths => [1, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30]
  REPORT_FIR_DATA[:second].each do |item|
    ws.add_row [nil, item[:週], item[:imp], item[:click], item[:CTR], item[:CPC], item[:CV], item[:CPA], item[:広告費], item[:総売上], item[:ROAS]], :height => 60, :style => [nil, st_date, st_number, st_number, st_percent, st_currency, st_number, st_currency, st_currency, st_currency, st_percent], :width => :auto
  end

  ws.add_conditional_formatting("I57:I71", { :type => :dataBar, :dxfId => st_profitable, :priority => 0, :data_bar => Axlsx::DataBar.new })
  ws.add_conditional_formatting("J57:J71", { :type => :dataBar, :dxfId => st_profitable, :priority => 1, :data_bar => Axlsx::DataBar.new })
  ws.add_conditional_formatting("K57:K71", { :type => :dataBar, :dxfId => st_profitable, :priority => 1, :data_bar => Axlsx::DataBar.new })

  ws.add_chart(Axlsx::CombineChart, :title => ' ', :bar_dir => :col) do |chart|
    chart.start_at 1, 48
    chart.end_at 11, 53
    chart.add_series 'bar', :data => ws["J57:J71"], :labels => ws["B57:B71"], :title => ws["J56"], :colors => (1..15).map{Constants::BLUE_COLOR}, :on_primary_axis => false
    chart.add_series 'line', :data => ws["K57:K71"], :labels => ws["B57:B71"], :title => ws["K56"], :color => Constants::YELLOW_COLOR, :show_marker => true
    chart.catAxis.label_rotation = -45
    chart.d_lbls.d_lbl_pos = :t
    chart.d_lbls.show_val = true
    chart.valAxis.gridlines = false
    chart.catAxis.gridlines = false
    chart.val_axis.format_code = '¥#,###,##0'
  end

  ws.add_row
  ws.add_row [nil, Constants::FSSHEET_COMMENT_TITLE], :style => st_tb_head, :height => 30
  ws.merge_cells "B73:K73"

  ws.add_row [nil, Constants::FSSHEET_LIST]
  ws.merge_cells "B74:K79"
  7.times{ ws.add_row }
  # define chart end

  # define chart start
  ws.add_row [nil, Constants::FSSHEET_SIX_TITLE], :style => st_title, :height => 30
  ws.add_row [nil, Constants::FSSHEET_CHART_TITLE], :style => st_tb_head, :height => 30
  ws.merge_cells "B83:K83"

  5.times{ ws.add_row [], :height => 90 }
  ws.add_row

  ws.add_row [nil, Constants::FSSHEET_SEV_TITLE], :style => st_title, :height => 30
  ws.add_row Constants::FSSHEET_STB_TITLES, :style => st_tb_head, :height => 30, :widths => [1, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30]
  REPORT_FIR_DATA[:third].each do |item|
    ws.add_row [nil, item[:週], item[:imp], item[:click], item[:CTR], item[:CPC], item[:CV], item[:CPA], item[:広告費], item[:総売上], item[:ROAS]], :height => 60, :style => [nil, st_date, st_number, st_number, st_percent, st_currency, st_number, st_currency, st_currency, st_currency, st_percent], :width => :auto
  end

  ws.add_conditional_formatting("I92:I106", { :type => :dataBar, :dxfId => st_profitable, :priority => 0, :data_bar => Axlsx::DataBar.new })
  ws.add_conditional_formatting("J92:J106", { :type => :dataBar, :dxfId => st_profitable, :priority => 1, :data_bar => Axlsx::DataBar.new })
  ws.add_conditional_formatting("K92:K106", { :type => :dataBar, :dxfId => st_profitable, :priority => 1, :data_bar => Axlsx::DataBar.new })

  ws.add_chart(Axlsx::CombineChart, :title => ' ', :bar_dir => :col) do |chart|
    chart.start_at 1, 83
    chart.end_at 11, 88
    chart.add_series 'bar', :data => ws["J92:J106"], :labels => ws["B92:B106"], :title => ws["J91"], :colors => (1..15).map{Constants::BLUE_COLOR}, :on_primary_axis => false
    chart.add_series 'line', :data => ws["K92:K106"], :labels => ws["B92:B106"], :title => ws["K91"], :color => Constants::YELLOW_COLOR, :show_marker => true
    chart.catAxis.label_rotation = -45
    chart.d_lbls.d_lbl_pos = :t
    chart.d_lbls.show_val = true
    chart.valAxis.gridlines = false
    chart.catAxis.gridlines = false
    chart.val_axis.format_code = '¥#,###,##0'
  end

  ws.add_row
  ws.add_row [nil, Constants::FSSHEET_COMMENT_TITLE], :style => st_tb_head, :height => 30
  ws.merge_cells "B108:K108"

  ws.add_row [nil, Constants::FSSHEET_LIST]
  ws.merge_cells "B109:K114"
  7.times{ ws.add_row }
  # define chart end

  ws.add_row [nil, Constants::FSSHEET_SEC_SUMMARY], :style => st_summary, :height => 30
  ws.merge_cells "B117:K117"
  ws.add_row

  # define chart start
  ws.add_row [nil, Constants::FSSHEET_EGT_TITLE], :style => st_title, :height => 30
  ws.add_row [nil, Constants::FSSHEET_CHART_TITLE], :style => st_tb_head, :height => 30
  ws.merge_cells "B120:K120"

  5.times{ ws.add_row [], :height => 90 }
  ws.add_row

  ws.add_row [nil, Constants::FSSHEET_NIN_TITLE], :style => st_title, :height => 30
  ws.add_row Constants::FSSHEET_STB_TITLES, :style => st_tb_head, :height => 30, :widths => [1, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30]
  REPORT_FIR_DATA[:fourth].each do |item|
    ws.add_row [nil, item[:週], item[:imp], item[:click], item[:CTR], item[:CPC], item[:CV], item[:CPA], item[:広告費], item[:総売上], item[:ROAS]], :height => 60, :style => [nil, st_date, st_number, st_number, st_percent, st_currency, st_number, st_currency, st_currency, st_currency, st_percent], :width => :auto
  end

  ws.add_conditional_formatting("I129:I143", { :type => :dataBar, :dxfId => st_profitable, :priority => 0, :data_bar => Axlsx::DataBar.new })
  ws.add_conditional_formatting("J129:J143", { :type => :dataBar, :dxfId => st_profitable, :priority => 1, :data_bar => Axlsx::DataBar.new })
  ws.add_conditional_formatting("K129:K143", { :type => :dataBar, :dxfId => st_profitable, :priority => 1, :data_bar => Axlsx::DataBar.new })

  ws.add_chart(Axlsx::CombineChart, :title => " ", :bar_dir => :col) do |chart|
    chart.start_at 1, 120
    chart.end_at 11, 125
    chart.add_series 'bar', :data => ws["J129:J143"], :labels => ws["B129:B143"], :title => ws["J128"], :colors => (1..15).map{Constants::BLUE_COLOR}, :on_primary_axis => false
    chart.add_series 'line', :data => ws["K129:K143"], :labels => ws["B129:B143"], :title => ws["K128"], :color => Constants::YELLOW_COLOR, :show_marker => true
    chart.catAxis.label_rotation = -45
    chart.d_lbls.d_lbl_pos = :t
    chart.d_lbls.show_val = true
    chart.valAxis.gridlines = false
    chart.catAxis.gridlines = false
    chart.val_axis.format_code = '¥#,###,##0'
  end

  ws.add_row
  ws.add_row [nil, Constants::FSSHEET_COMMENT_TITLE], :style => st_tb_head, :height => 30
  ws.merge_cells "B145:K145"

  ws.add_row [nil, Constants::FSSHEET_LIST]
  ws.merge_cells "B146:K151"
  7.times{ ws.add_row }
  # define chart end

  # define chart start
  ws.add_row [nil, Constants::FSSHEET_TEN_TITLE], :style => st_title, :height => 30
  ws.add_row [nil, Constants::FSSHEET_CHART_TITLE], :style => st_tb_head, :height => 30
  ws.merge_cells "B155:K155"

  5.times{ ws.add_row [], :height => 90 }
  ws.add_row

  ws.add_row [nil, Constants::FSSHEET_TFIR_TITLE], :style => st_title, :height => 30
  ws.add_row Constants::FSSHEET_STB_TITLES, :style => st_tb_head, :height => 30, :widths => [1, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30]
  REPORT_FIR_DATA[:fifth].each do |item|
    ws.add_row [nil, item[:週], item[:imp], item[:click], item[:CTR], item[:CPC], item[:CV], item[:CPA], item[:広告費], item[:総売上], item[:ROAS]], :height => 60, :style => [nil, st_date, st_number, st_number, st_percent, st_currency, st_number, st_currency, st_currency, st_currency, st_percent], :width => :auto
  end

  ws.add_conditional_formatting("I164:I178", { :type => :dataBar, :dxfId => st_profitable, :priority => 0, :data_bar => Axlsx::DataBar.new })
  ws.add_conditional_formatting("J164:J178", { :type => :dataBar, :dxfId => st_profitable, :priority => 1, :data_bar => Axlsx::DataBar.new })
  ws.add_conditional_formatting("K164:K178", { :type => :dataBar, :dxfId => st_profitable, :priority => 1, :data_bar => Axlsx::DataBar.new })

  ws.add_chart(Axlsx::CombineChart, :title => " ", :bar_dir => :col) do |chart|
    chart.start_at 1, 155
    chart.end_at 11, 160
    chart.add_series 'bar', :data => ws["J164:J178"], :labels => ws["B164:B178"], :title => ws["J163"], :colors => (1..15).map{Constants::BLUE_COLOR}, :on_primary_axis => false
    chart.add_series 'line', :data => ws["K164:K178"], :labels => ws["B164:B178"], :title => ws["K163"], :color => Constants::YELLOW_COLOR, :show_marker => true
    chart.catAxis.label_rotation = -45
    chart.d_lbls.d_lbl_pos = :t
    chart.d_lbls.show_val = true
    chart.valAxis.gridlines = false
    chart.catAxis.gridlines = false
    chart.val_axis.format_code = '¥#,###,##0'
  end


  # define chart start
  ws.add_row
  ws.add_row [nil, Constants::FSSHEET_COMMENT_TITLE], :style => st_tb_head, :height => 30
  ws.merge_cells "B180:K180"

  ws.add_row [nil, Constants::FSSHEET_LIST]
  ws.merge_cells "B181:K186"
  7.times{ ws.add_row }

  ws.add_row [nil, Constants::FSSHEET_TSEC_TITLE], :style => st_title, :height => 30
  ws.add_row [nil, Constants::FSSHEET_CHART_TITLE], :style => st_tb_head, :height => 30
  ws.merge_cells "B190:K190"

  5.times{ ws.add_row [], :height => 90 }
  ws.add_row

  ws.add_row [nil, Constants::FSSHEET_TTHI_TITLE], :style => st_title, :height => 30
  ws.add_row Constants::FSSHEET_STB_TITLES, :style => st_tb_head, :height => 30, :widths => [1, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30]
  REPORT_FIR_DATA[:sixth].each do |item|
    ws.add_row [nil, item[:週], item[:imp], item[:click], item[:CTR], item[:CPC], item[:CV], item[:CPA], item[:広告費], item[:総売上], item[:ROAS]], :height => 60, :style => [nil, st_date, st_number, st_number, st_percent, st_currency, st_number, st_currency, st_currency, st_currency, st_percent], :width => :auto
  end

  ws.add_conditional_formatting("I199:I213", { :type => :dataBar, :dxfId => st_profitable, :priority => 0, :data_bar => Axlsx::DataBar.new })
  ws.add_conditional_formatting("J199:J213", { :type => :dataBar, :dxfId => st_profitable, :priority => 1, :data_bar => Axlsx::DataBar.new })
  ws.add_conditional_formatting("K199:K213", { :type => :dataBar, :dxfId => st_profitable, :priority => 1, :data_bar => Axlsx::DataBar.new })

  ws.add_chart(Axlsx::CombineChart, :title => " ", :bar_dir => :col) do |chart|
    chart.start_at 1, 190
    chart.end_at 11, 195
    chart.add_series 'bar', :data => ws["J199:J213"], :labels => ws["B199:B213"], :title => ws["J198"], :colors => (1..15).map{Constants::BLUE_COLOR}, :on_primary_axis => false
    chart.add_series 'line', :data => ws["K199:K213"], :labels => ws["B199:B213"], :title => ws["K198"], :color => Constants::YELLOW_COLOR, :show_marker => true
    chart.catAxis.label_rotation = -45
    chart.d_lbls.d_lbl_pos = :t
    chart.d_lbls.show_val = true
    chart.valAxis.gridlines = false
    chart.catAxis.gridlines = false
    chart.val_axis.format_code = '¥#,###,##0'
  end

  #define chart start
  ws.add_row
  ws.add_row [nil, Constants::FSSHEET_COMMENT_TITLE], :style => st_tb_head, :height => 30
  ws.merge_cells "B215:K215"

  ws.add_row [nil, Constants::FSSHEET_LIST]
  ws.merge_cells "B216:K221"
  7.times{ ws.add_row }

  # page break
  ws.page_setup.fit_to :width => 1, :height => 4
  ws.sheet_view.view = :page_break_preview # so you can see the breaks!

  #footer
  ws.add_row
  ws.add_row [nil, WB_FOOTER_TITLE], :style => st_footer
  ws.merge_cells "B225:K225"

end

# define your second sheet
wb.add_worksheet(:name => Constants::SHEET_SEC_NAME) do |ws|
  ws.add_row
  ws.add_row [nil, WB_HEAD_TITLE], :style => st_header, :width => :auto, :height => 40

  ws.add_row [nil, WB_FS_SUBHEAD_TITLE], :style => st_sub_header, :height => 30, :width => :auto
  ws.add_row [nil, WB_SE_SUBHEAD_TITLE], :style => st_sub_header, :height => 30, :width => :auto

  ws.add_row
  ws.add_row Constants::SCSHEET_TB_TITLES, :style => st_tb_head, :height => 30, :widths => [1, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30]

  for sec_index in 6..67
    ws.merge_cells "D#{sec_index}:E#{sec_index}"
  end

  SEC_REPORT_DATA.each do |item|
    if item.nil?
      ws.add_row
    end
    ws.add_row [nil, item[:広告費投下比率], item[:No], item[:商品カテゴリ名], nil, item[:広告費], item[:imp], item[:click], item[:CTR], item[:CPC], item[:CV], item[:CPA], item[:全体総売上], item[:ROAS]], :height => 30, :style => [nil, st_number, st_number, st_tb_body, nil, st_currency, st_number, st_number, st_percent, st_currency, st_number, st_currency, st_currency, st_percent], :width => :auto
  end

  # page break
  ws.page_setup.fit_to :width => 1, :height => 1
  ws.sheet_view.view = :page_break_preview # so you can see the breaks!

  #footer
  4.times { ws.add_row }
  ws.add_row [nil, WB_FOOTER_TITLE], :style => st_footer
  ws.merge_cells "B70:N70"

end


#define your third sheet
wb.add_worksheet(:name => Constants::SHEET_THI_NAME) do |ws|
  ws.add_row
  ws.add_row [nil, WB_HEAD_TITLE], :style => st_header, :width => :auto, :height => 40

  ws.add_row [nil, WB_FS_SUBHEAD_TITLE], :style => st_sub_header, :height => 30, :width => :auto
  ws.add_row [nil, WB_SE_SUBHEAD_TITLE], :style => st_sub_header, :height => 30, :width => :auto

  2.times { ws.add_row }
  ws.add_row Constants::THSHEET_TB_TITLES, :style => st_tb_head, :height => 30, :widths => [1, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30]

  for thi_index in 7..12
    ws.merge_cells "D#{thi_index}:E#{thi_index}"
  end

  THI_REPORT_DATA.each do |item|
    ws.add_row [nil, item[:広告費投下比率], item[:No], item[:キャンペーン名], nil, item[:広告費], item[:imp], item[:click], item[:CTR], item[:CPC], item[:CV], item[:CPA], item[:全体総売上], item[:ROAS]], :height => 30, :style => [nil, st_number, st_number, st_tb_body, nil, st_currency, st_number, st_number, st_percent, st_currency, st_number, st_currency, st_currency, st_percent], :width => :auto
  end

  # page break
  ws.page_setup.fit_to :width => 1, :height => 1
  ws.sheet_view.view = :page_break_preview # so you can see the breaks!

  #footer
  3.times { ws.add_row }
  ws.add_row [nil, WB_FOOTER_TITLE], :style => st_footer
  ws.merge_cells "B16:N16"

end

#define your fourth sheet
wb.add_worksheet(:name => Constants::SHEET_FOR_NAME) do |ws|
  ws.add_row
  ws.add_row [nil, WB_HEAD_TITLE], :style => st_header, :width => :auto, :height => 40

  ws.add_row [nil, WB_FS_SUBHEAD_TITLE], :style => st_sub_header, :height => 30, :width => :auto
  ws.add_row [nil, WB_SE_SUBHEAD_TITLE], :style => st_sub_header, :height => 30, :width => :auto

  ws.add_row
  ws.add_row Constants::FOSHEET_TB_TITLES, :style => st_tb_head, :height => 30, :widths => [1, 10, 60, 60, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30]
  FOR_REPORT_DATA.each do |item|
    ws.add_row [nil, item[:No], item[:キャンペーン名], item[:キーワード], item[:マッチタイプ], item[:キーワード種別], item[:広告費], item[:imp], item[:click], item[:CTR], item[:CPC], item[:CV], item[:CPA], item[:全体総売上], item[:ROAS]], :height => 30, :style => [nil, st_number, st_tb_body, st_tb_body, st_tb_body, st_tb_body, st_currency, st_number, st_number, st_percent, st_currency, st_number, st_currency, st_currency, st_percent], :width => :auto
  end

  # page break
  ws.page_setup.fit_to :width => 1, :height => 1
  ws.sheet_view.view = :page_break_preview # so you can see the breaks!

  #footer
  2.times { ws.add_row }
  ws.add_row [nil, WB_FOOTER_TITLE], :style => st_footer
  ws.merge_cells "B14:O14"

end

#define your fifth sheet
wb.add_worksheet(:name => Constants::SHEET_FIF_NAME) do |ws|
  ws.add_row
  ws.add_row [nil, WB_HEAD_TITLE], :style => st_header, :width => :auto, :height => 40

  ws.add_row [nil, WB_FS_SUBHEAD_TITLE], :style => st_sub_header, :height => 30, :width => :auto
  ws.add_row [nil, WB_SE_SUBHEAD_TITLE], :style => st_sub_header, :height => 30, :width => :auto

  ws.add_row
  ws.add_row Constants::FISHEET_TB_TITLES, :style => st_tb_head, :height => 30, :widths => [1, 10, 60, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30]
  FIF_REPORT_DATA.each do |item|
    ws.add_row [nil, item[:No], item[:商品名], item[:ASIN], item[:広告費], item[:imp], item[:click], item[:CTR], item[:CPC], item[:CV], item[:CPA], item[:全体総売上], item[:ROAS]], :height => 30, :style => [nil, st_number, st_tb_body, st_tb_body, st_currency, st_number, st_number, st_percent, st_currency, st_number, st_currency, st_currency, st_percent], :width => :auto
  end

  # page break
  ws.page_setup.fit_to :width => 1, :height => 1
  ws.sheet_view.view = :page_break_preview # so you can see the breaks!

  #footer
  2.times { ws.add_row }
  ws.add_row [nil, WB_FOOTER_TITLE], :style => st_footer
  ws.merge_cells "B14:M14"

end


p.serialize('月次レポート 依頼用.xlsx')