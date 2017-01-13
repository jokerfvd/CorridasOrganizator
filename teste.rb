# encoding: UTF-8
require 'rubygems'
require 'watir'
require 'time'
require 'nokogiri'
require 'open-uri'

def englishMonth(data)
	begin
		data['Fev'] = 'Feb'
		data['Abr'] = 'Apr'
		data['Mai'] = 'May'
		data['Ago'] = 'Aug'
		data['Set'] = 'Sep'
		data['Out'] = 'Oct'	
		data['Dez'] = 'Dec'	
	rescue
	end
	return data
end

def getIdDaCorrida(url)
	aux = url.split("/")
	aux.size.times{|i|
		if (aux[i] == "corrida-de-rua")
			return aux[i+1]
		end
	}
	return nil
end


#=begin
#ruby teste.rb "Rio de Janeiro;Niterói;Petrópolis;São Gonçalo"
locais = ARGV[0].encode('utf-8').split(";")

browser = Watir::Browser.new 
browser.goto "http://www.ativo.com/calendario/"
browser.select_list(:id => "modalidade_select").select("Corrida de Rua")
articles = browser.elements(:xpath => "//*[@id='container']/article")

vetorArticles = []
articles.each do |article|
	vetorArticles << article
end

vetorArticles.each do |article|
	local = article.element(:xpath => "div[1]/div[1]/header/a/div[4]/div/span").text
	if (locais.include? local)
		nome = article.element(:xpath => "div[1]/div[1]/header/a/div[2]/h2").text
		aux = article.element(:xpath => "div[1]/div[1]/header/a/time")
		data = Time.parse(englishMonth("#{aux.spans[1].text}/#{aux.spans[2].text}/#{aux.spans[0].text}"))
		id = getIdDaCorrida(article.element(:xpath => "div[1]/figure/a").attribute_value('href'))
		
		puts "#{data.strftime("%d/%m/%Y")} -- #{local} -- #{id} --> #{nome}"
		
		#indo para o link da corrida
		browser2 = Watir::Browser.new 
		browser2.goto "https://checkout.ativo.com/evento/"+id
		lista = browser2.element(:xpath => "/html/body/main/div/div/div[1]/ul").lis
		vetorLista = []
		lista.each do |li|
			puts li.text
			vetorLista << li
		end
		puts "1111111--#{lista.size}--#{vetorLista.size}"
		vetorLista.each do |li|
			descricao = li.element(:class => "descricao").spans[0].text.strip.chomp
			puts "\t"+descricao
			radios = li.elements(:name => "modalidade")
			radios.each do |radio|
				parent = radio.parent
				modalidade = nil
				if parent.labels[0].exists?
					modalidade = parent.labels[0].text
				elsif 	parent.spans[0].exists?
					modalidade = parent.spans[0].text
				end
				puts "\t\t"+modalidade
			end
			puts "222222"
		end
		browser2.close
		
	end
end


#=end
