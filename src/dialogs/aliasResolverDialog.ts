import { Attachment, CardFactory } from 'botbuilder';
import {
  DialogTurnResult,
  OAuthPrompt,
  PromptValidatorContext,
  TextPrompt,
  WaterfallDialog,
  WaterfallStepContext
} from 'botbuilder-dialogs';
import { GraphHelper } from '../helpers/graphHelper';
import GenericCard from '../resources/generic.json';
import { HelperDialog } from './helperDialog';

const TEXT_PROMPT = 'textPrompt';
const WATERFALL_DIALOG = 'waterfallDialog';
const OAUTH_PROMPT = 'OAuthPrompt';

export class AliasResolverDialog extends HelperDialog {
  private static tokenResponse: any;
  
  private static async aliasPromptValidator(promptContext: PromptValidatorContext<string>): Promise<boolean> {
    if (promptContext.recognized.succeeded) {
      
      const alias: string = promptContext.recognized.value;
      
      if (await GraphHelper.aliasExists(AliasResolverDialog.tokenResponse, alias))  {
        
        const aliasExistText = 'Alias already exist';
        const aliasExistCard: Attachment = CardFactory.adaptiveCard(JSON.parse(
          JSON.stringify(GenericCard).replace('@{placeholder}', aliasExistText)));
    
        await promptContext.context.sendActivity({ attachments: [aliasExistCard] });
        
        return false;
      }

      return true;

    } else {
      return false;
    }
  }

  constructor(id: string) {
    super(id || 'aliasResolverDialog', process.env.connectionName);
    
    this
        .addDialog(new TextPrompt(TEXT_PROMPT, AliasResolverDialog.aliasPromptValidator.bind(this)))
        .addDialog(new OAuthPrompt(OAUTH_PROMPT, {
          connectionName: process.env.connectionName,
          text: 'Please Sign In',
          timeout: 300000,
          title: 'Sign In'
        }))
        .addDialog(new WaterfallDialog(WATERFALL_DIALOG, [
          this.promptStep.bind(this),
          this.initialStep.bind(this),
          this.finalStep.bind(this)
        ]));

    this.initialDialogId = WATERFALL_DIALOG;

  }

  /**
   * Prompt step in the waterfall. 
   */
  private async promptStep(stepContext: WaterfallStepContext): Promise<DialogTurnResult> {
      return await stepContext.beginDialog(OAUTH_PROMPT);
  }

  private async initialStep(stepContext: WaterfallStepContext): Promise<DialogTurnResult> {
    const tokenResponse = stepContext.result;
    if (tokenResponse && tokenResponse.token) {
      
      AliasResolverDialog.tokenResponse = tokenResponse;

      const siteDetails = (stepContext.options as any).siteDetails;
      const promptMsg = `Provide an alias for your ${siteDetails.siteType} site`;

      if (!siteDetails.alias) {

        const aliasCard: Attachment = CardFactory.adaptiveCard(JSON.parse(
          JSON.stringify(GenericCard).replace('@{placeholder}', promptMsg)));
  
        return await stepContext.prompt(TEXT_PROMPT, {prompt: { attachments: [aliasCard] }});
      } else {
        return await stepContext.next(siteDetails.alias);
      }
    }
    await stepContext.context.sendActivity('Login was not successful please try again.');
    return await stepContext.endDialog();
  }

  private async finalStep(stepContext: WaterfallStepContext): Promise<DialogTurnResult> {    
    const owner = stepContext.result;
    return await stepContext.endDialog(owner);
  }
}
