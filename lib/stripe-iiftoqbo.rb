require 'csv'
require_relative 'iif'
require_relative 'ofx'

module StripeIIFToQBO
  class Converter

    MAX_LINES = 500

    def initialize(options={})
      @account_id = options[:account_id] if options[:account_id]
      @iif_file = options[:iif_file] if options[:iif_file]
      @payments_file = options[:payments_file] if options[:payments_file]
      @transfers_file = options[:transfers_file] if options[:transfers_file]
      @server_time = options[:server_time] || Date.today
      @output_file = options[:output_file] if options[:output_file]
      raise 'missing required iif file' if @iif_file.nil?
      raise 'missing required output file' if @output_file.nil?
      load_payments_file(@payments_file)
      load_transfers_file(@transfers_file)
      load_iif_file(@iif_file)
    end

    def load_payments_file(payments_file)
      @payments = {}

      if payments_file
		CSV.foreach(payments_file, :headers => true, :encoding => 'windows-1251:utf-8') do |row|
			# when exporting CSV from https://dashboard.stripe.com/payments (unified payments)
			if row['id']
		  		@payments[row['id']] = "[" + ( row['Currency'] || '???' ) + "] " + ( row['Description'] || '' ) + " {" + ( row['Card Name'] || '' )  + "} " + ( row['Customer Email'] || '' ) + " | " + ( row['Customer ID'] || '' ) + " | " + ( row['Card Address State'] || '' )
			end
			# when exporting CSV from https://dashboard.stripe.com/balance
			if row['Source']
				@payments[row['Source']] = "[" + ( row['Currency'] || '???' ) + "] " + ( row['Description'] || '' ) + " | " + ( row['id'] || '' ) + " | "
			end
		end
      end
    end

    def load_transfers_file(transfers_file)
      @transfers = {}

      if transfers_file
		CSV.foreach(transfers_file, :headers => true, :encoding => 'windows-1251:utf-8') do |row|
			# when exporting CSV from https://dashboard.stripe.com/payouts
			if row['id']
		  		@transfers[row['id']] = "[" + ( row['Currency'] || '???' ) + "] " + " | " + ( row['Destination'] || '' ) + " | " + ( row['Balance Transaction'] || '' )
			end
			# when exporting CSV from https://dashboard.stripe.com/balance
			if row['Source']
				@transfers[row['Source']] = "[" + ( row['Currency'] || '???' ) + "] " + ( row['id'] || '' )
			end
			# default behavior, unknown use
			if row['ID']
				@transfers[row['ID']] = row['Description'] || ''
			end
        end
      end
    end

    def load_iif_file(iif_file)
      @ofx_entries = []
      file_count = 0
      if iif_file
        IIF(iif_file) do |iif|
          iif.transactions.each do |transaction|
            #process the transaction
            transaction.entries.each do |iif_entry|
              ofx_entry = convert_iif_entry_to_ofx(iif_entry)
              if ofx_entry
                @ofx_entries.push(ofx_entry)
              end
            end
            #write file (if necessary)
            if @ofx_entries.length >= MAX_LINES
              write_qbo_file(file_count)
              file_count += 1
              @ofx_entries = []
			end
          end
        end
        #write file (anything left)
        write_qbo_file(file_count)
      end
    end

    def convert_iif_entry_to_ofx(iif_entry)
      ofx_entry = {}
      ofx_entry[:date] = iif_entry.date
      ofx_entry[:fitid] = iif_entry.memo
      ofx_entry[:accnt] = iif_entry.accnt
      ofx_entry[:trnstype] = iif_entry.trnstype
      ofx_entry[:memo] = iif_entry.memo
      ofx_entry[:currency] = "usd"
      case iif_entry.accnt
        when 'Stripe Third-party Account'
          ofx_entry[:amount] = -iif_entry.amount
          ofx_entry[:name] = iif_entry.name
          ofx_entry[:memo] =~ /Transfer from Stripe: (\S+)/
          transfer_id = $1
          if @transfers[transfer_id]
            ofx_entry[:memo] = "#{@transfers[transfer_id]} | #{iif_entry.memo}"
          end
		when 'Stripe Checking Account'
          ofx_entry[:memo] =~ /Transfer ID: (\S+)/
		  transfer_id = $1
		  ofx_entry[:trnstype] = "XFER"
          ofx_entry[:amount] = -iif_entry.amount
		  ofx_entry[:name] = "Transfer to #{iif_entry.accnt}"
		  ofx_entry[:currency] = "usd"
		  if @transfers[transfer_id]
			ofx_entry[:memo] = "#{@transfers[transfer_id]} | #{iif_entry.memo}"
			ofx_entry[:currency] = @transfers[transfer_id].split("[").last.split("]").first
		  end
		when 'Stripe Payment Processing Fees'
          ofx_entry[:memo] =~ /Fees for charge ID: (\S+)/
		  charge_id = $1
          ofx_entry[:amount] = -iif_entry.amount
		  ofx_entry[:name] = 'Stripe'
		  ofx_entry[:trnstype] = "FEE"
		  ofx_entry[:currency] = "usd"
          if @payments[charge_id]
            ofx_entry[:memo] = ofx_entry[:memo] + " | Processing Fees \n " + "#{@payments[charge_id]}"
			ofx_entry[:fitid] = charge_id
			ofx_entry[:currency] = @payments[charge_id].split("[").last.split("]").first
		    ofx_entry[:name] = "Stripe (#{ofx_entry[:currency].upcase})"

          end
        when 'Stripe Sales'
          ofx_entry[:amount] = -iif_entry.amount
          if iif_entry.memo =~ /Stripe Connect fee/
            ofx_entry[:name] = 'Stripe Connect Charge'
          elsif iif_entry.memo =~ /Charge/
            ofx_entry[:name] = 'Credit Card Charge'
          else
            ofx_entry[:name] = iif_entry.accnt
          end
          ofx_entry[:memo] =~ /Charge ID: (\S+)/
		  charge_id = $1
		  ofx_entry[:currency] = "usd"
          if @payments[charge_id]
            ofx_entry[:memo] = "#{@payments[charge_id]}" + " \n " + ofx_entry[:memo] 
			ofx_entry[:fitid] = charge_id
			ofx_entry[:currency] = @payments[charge_id].split("[").last.split("]").first
            ofx_entry[:name] = @payments[charge_id].split("{").last.split("}").first

          end
        when 'Stripe Returns'
          ofx_entry[:amount] = -iif_entry.amount
          ofx_entry[:name] = 'Credit Card Refund'
		  ofx_entry[:currency] = "usd"
          ofx_entry[:memo] =~ /Refund of charge (\S+)/
          charge_id = $1

          if @payments[charge_id]
			ofx_entry[:memo] = "#{@payments[charge_id]} Refund of Charge ID: #{charge_id}"
			ofx_entry[:currency] = @payments[charge_id].split("[").last.split("]").first
          end
        when 'Stripe Other Income'
          ofx_entry[:amount] = -iif_entry.amount
          ofx_entry[:name] = 'Other Income'
        when 'Stripe Account'
          #unnecessary
          return nil
      end
      if ofx_entry[:amount] == BigDecimal(0)
        #bail on zero amounts
        return nil
      end
      ofx_entry
    end

    def to_csv
      rows = []
      rows.push(['Date', 'Name', 'Account', 'Memo', 'Amount'].to_csv)
      @ofx_entries.each do |ofx_entry|
        rows.push([ofx_entry[:date].strftime('%m/%d/%Y'), ofx_entry[:name], ofx_entry[:accnt], "#{ofx_entry[:trnstype]} #{ofx_entry[:memo]}", ofx_entry[:amount].to_s('F')].to_csv)
      end
      return rows.join
    end

    def write_qbo_file(file_number)
      file_name = "#{@output_file}_#{file_number}.qbo"
      File.open(file_name, 'w') {|file| file.write to_qbo.to_s}
    end

    def to_qbo
      min_date = nil
      max_date = nil
      @ofx_entries.each do |e|
        if e[:date]
          min_date = e[:date] if min_date.nil? or e[:date] < min_date
          max_date = e[:date] if max_date.nil? or e[:date] > max_date
        end
      end
      ofx_builder = OFX::Builder.new do |ofx|
        ofx.dtserver = @server_time
        ofx.fi_org = 'Stripe'
        ofx.fi_fid = '0'
        ofx.bank_id = '123456789'
        ofx.acct_id = @account_id
        ofx.acct_type = 'CHECKING'
        ofx.dtstart = min_date
        ofx.dtend = max_date
        ofx.bal_amt = 0
        ofx.dtasof = max_date
      end

	  @ofx_entries.each do |ofx_entry|
        ofx_builder.transaction do |ofx_tr|
          ofx_tr.dtposted = ofx_entry[:date]
          ofx_tr.trntype = ofx_entry[:trnstype]
          ofx_tr.currency = ofx_entry[:currency]
          ofx_tr.trnamt = ofx_entry[:amount]
          ofx_tr.fitid = ofx_entry[:fitid]
          ofx_tr.name = ofx_entry[:name]
          ofx_tr.memo = ofx_entry[:memo]
        end
      end
      return ofx_builder.to_ofx
    end

  end
end
