Rails.application.routes.draw do
  get '/emails', to: 'emails#show'
end
