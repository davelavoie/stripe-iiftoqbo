module OFX
  class Transaction
    attr_accessor :trntype
    attr_accessor :currency
    attr_accessor :dtposted
    attr_accessor :trnamt
    attr_accessor :fitid
    attr_accessor :name
    attr_accessor :memo

    def trnamt=(amt)
      @trnamt = BigDecimal(amt.to_s)
    end
  end
end


