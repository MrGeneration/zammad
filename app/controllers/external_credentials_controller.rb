# Copyright (C) 2012-2026 Zammad Foundation, https://zammad-foundation.org/

class ExternalCredentialsController < ApplicationController
  include ExternalCredential::SensitiveAttributes

  prepend_before_action :authenticate_and_authorize!

  def index
    model_index_render(ExternalCredential, params)
  end

  def show
    model_show_render(ExternalCredential, params)
  end

  def create
    model_create_render(ExternalCredential, params)
  end

  def update
    model_update_render(ExternalCredential, params)
  end

  def destroy
    model_destroy_render(ExternalCredential, params)
  end

  def app_verify
    render json: { attributes: ExternalCredential.app_verify(params.permit!.to_h) }, status: :ok
  rescue => e
    logger.error e
    render json: { error: e.message }, status: :ok
  end

  def link_account
    provider     = params[:provider].downcase
    notification = [true, 'true'].include?(params[:notification])
    ensure_notification_setup_allowed! if notification

    attributes = ExternalCredential.request_account_to_link(provider, params)
    session[:request_token] = attributes[:request_token]
    session[:code_verifier] = attributes[:code_verifier]
    session[:channel_id] = params[:channel_id]
    session[:shared_mailbox] = params[:shared_mailbox]
    session[:notification] = notification
    request.session_options.delete(:skip)
    redirect_to attributes[:authorize_url], allow_other_host: true
  end

  def callback
    provider = params[:provider].downcase

    if session[:notification] && provider == 'microsoft_graph'
      return handle_notification_callback
    end

    channel = ExternalCredential.link_account(provider, session[:request_token], link_params)
    return redirect_to(channel), allow_other_host: true if channel.instance_of?(String)

    redirect_to app_url(provider, channel.id), allow_other_host: true
  ensure
    session[:request_token]  = nil
    session[:code_verifier]  = nil
    session[:channel_id]     = nil
    session[:shared_mailbox] = nil
    session[:notification]   = nil
  end

  private

  # System notifications of hosted instances are sent by the hosting provider.
  def ensure_notification_setup_allowed!
    raise Exceptions::Forbidden if Setting.get('system_online_service')
  end

  def handle_notification_callback
    ensure_notification_setup_allowed!

    # Return to the admin interface if admin consent is in process and the user clicks on "Back to app".
    return redirect_to(notification_error_url('AADSTS65004'), allow_other_host: true) if params[:error_description].to_s.include?('AADSTS65004')

    auth_data     = fetch_graph_auth_data(notification_external_credential)
    configuration = {
      user:           extract_preferred_username(auth_data[:id_token]),
      shared_mailbox: session[:shared_mailbox].presence,
    }

    # Only switch the notification delivery if the mailbox accepts mails from the notification sender.
    probe_result = probe_notification_delivery(configuration, auth_data)
    if probe_result[:result] != 'ok'
      return redirect_to(notification_error_url('notification_verification_failed', probe_result[:message]), allow_other_host: true)
    end

    Service::System::SetEmailNotificationConfiguration.execute(
      adapter:              'microsoft_graph_outbound',
      new_configuration:    configuration,
      microsoft_graph_auth: auth_data,
    )

    redirect_to channels_email_url, allow_other_host: true
  end

  def notification_external_credential
    raise Exceptions::UnprocessableContent, __('Invalid OAuth state parameter.') if params[:state] != session[:request_token]

    external_credential = ExternalCredential.find_by(name: 'microsoft_graph')
    raise Exceptions::UnprocessableContent, __('No Microsoft Graph app configured!') if !external_credential
    raise Exceptions::UnprocessableContent, __("The required parameter 'code' is missing.") if params[:code].blank?

    external_credential
  end

  def probe_notification_delivery(configuration, auth_data)
    EmailHelper::Probe.outbound(
      {
        adapter: 'microsoft_graph_outbound',
        options: configuration.merge(password: auth_data[:access_token]).compact_blank,
      },
      Setting.get('notification_sender'),
    )
  end

  def notification_error_url(error_code, message = nil)
    url = "#{Setting.get('http_type')}://#{Setting.get('fqdn')}/#channels/microsoft_graph/error/#{error_code}"
    return url if message.blank?

    "#{url}/param/#{CGI.escapeURIComponent(message.truncate(300))}"
  end

  def fetch_graph_auth_data(external_credential)
    response = ExternalCredential::MicrosoftGraph.authorize_tokens(external_credential.credentials, params[:code])
    %w[refresh_token access_token expires_in scope token_type id_token].each do |key|
      raise Exceptions::UnprocessableContent, "No #{key} for authorization request found!" if response[key.to_sym].blank?
    end

    response.merge(
      provider:      'microsoft_graph',
      type:          'XOAUTH2',
      client_id:     external_credential.credentials[:client_id],
      client_secret: external_credential.credentials[:client_secret],
      client_tenant: external_credential.credentials[:client_tenant],
    )
  end

  def extract_preferred_username(id_token)
    user_data = ExternalCredential::MicrosoftGraph.user_info(id_token)
    raise Exceptions::UnprocessableContent, __("The user's 'preferred_username' could not be extracted from 'id_token'.") if user_data[:preferred_username].blank?

    user_data[:preferred_username]
  end

  def channels_email_url
    "#{Setting.get('http_type')}://#{Setting.get('fqdn')}/#channels/email"
  end

  def link_params
    params.permit!.to_h.merge(channel_id: session[:channel_id], shared_mailbox: session[:shared_mailbox], code_verifier: session[:code_verifier])
  end

  def callback_url(provider)
    ExternalCredential.callback_url(provider)
  end

  def app_url(provider, channel_id)
    ExternalCredential.app_url(provider, channel_id)
  end
end
