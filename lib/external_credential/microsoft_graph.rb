# Copyright (C) 2012-2026 Zammad Foundation, https://zammad-foundation.org/

class ExternalCredential::MicrosoftGraph < ExternalCredential::MicrosoftBase
  def self.channel_area
    'MicrosoftGraph::Account'.freeze
  end

  def self.error_missing_app_configuration
    __('No Microsoft Graph app configured!')
  end

  def self.authorize_scope
    'offline_access openid profile email mail.readwrite mail.readwrite.shared mail.send mail.send.shared'
  end

  # The email notification channel only sends, so it does not need access to the mailbox content.
  def self.notification_authorize_scope
    'offline_access openid profile email mail.send mail.send.shared'
  end

  def self.generate_authorize_url(credentials, scope: authorize_scope, state: nil, code_challenge: nil)
    return super if [true, 'true'].exclude?(credentials[:notification])

    super(credentials, scope: notification_authorize_scope, state:, code_challenge:)
  end

  # The email notification channel keeps its own copy of the client secret, outside of the channel area.
  def self.update_client_secret(previous_client_secret, current_client_secret)
    super

    Channel.in_area('Email::Notification').find_each do |channel|
      next if channel.options.dig(:auth, :provider) != provider_name
      next if channel.options.dig(:auth, :client_secret) != previous_client_secret

      channel.options[:auth][:client_secret] = current_client_secret
      channel.save!
    end

    true
  end

  def self.channel_options_inbound(user_data, account_data)
    {
      adapter: 'microsoft_graph_inbound',
      options: {
        user:           user_data[:preferred_username],
        shared_mailbox: account_data[:shared_mailbox],
      }.compact_blank,
    }
  end

  def self.channel_options_outbound(user_data, account_data)
    {
      adapter: 'microsoft_graph_outbound',
      options: {
        user:           user_data[:preferred_username],
        shared_mailbox: account_data[:shared_mailbox],
      }.compact_blank,
    }
  end
end
