package org.braisdom.excel.impl;

import com.helger.commons.ValueEnforcer;
import com.helger.commons.collection.impl.CommonsArrayList;
import com.helger.commons.collection.impl.ICommonsList;
import com.helger.css.ECSSUnit;
import com.helger.css.ECSSVersion;
import com.helger.css.decl.CSSDeclaration;
import com.helger.css.decl.CSSExpression;
import com.helger.css.decl.CSSExpressionMemberTermSimple;
import com.helger.css.decl.ICSSExpressionMember;
import com.helger.css.decl.shorthand.CSSPropertyWithDefaultValue;
import com.helger.css.decl.shorthand.CSSShortHandDescriptor;
import com.helger.css.property.CCSSProperties;
import com.helger.css.property.ICSSProperty;
import com.helger.css.writer.CSSWriterSettings;

import javax.annotation.Nonnull;

public class CSSShortHandDescriptorWrapper extends CSSShortHandDescriptor {

    private final ICommonsList<CSSPropertyWithDefaultValue> m_aSubProperties;

    public CSSShortHandDescriptorWrapper(CSSShortHandDescriptor cssShortHandDescriptor) {
        super(cssShortHandDescriptor.getProperty(), new CSSPropertyWithDefaultValue (CCSSProperties.BORDER_TOP_WIDTH,
                ECSSUnit.px (3)));
        // TODO
        m_aSubProperties = cssShortHandDescriptor.getAllSubProperties();
    }

    @Nonnull
    @Override
    public ICommonsList<CSSDeclaration> getSplitIntoPieces(@Nonnull CSSDeclaration aDeclaration) {
        ValueEnforcer.notNull(aDeclaration, "Declaration");

        // Check that declaration matches this property
        if (!aDeclaration.hasProperty(getProperty()))
            throw new IllegalArgumentException("Cannot split a '" +
                    aDeclaration.getProperty() +
                    "' as a '" +
                    getProperty().getName() +
                    "'");

        // global
        final int nSubProperties = m_aSubProperties.size();
        final ICommonsList<CSSDeclaration> ret = new CommonsArrayList<>();
        final ICommonsList<ICSSExpressionMember> aExpressionMembers = aDeclaration.getExpression().getAllMembers();

        // Modification for margin and padding
        modifyExpressionMembers(aExpressionMembers);
        final int nExpressionMembers = aExpressionMembers.size();
        final CSSWriterSettings aCWS = new CSSWriterSettings(ECSSVersion.CSS30, false);
        final boolean[] aHandledSubProperties = new boolean[aExpressionMembers.size()];

        // For all expression members
        for (int nExprMemberIndex = 0; nExprMemberIndex < nExpressionMembers; ++nExprMemberIndex) {
            final ICSSExpressionMember aMember = aExpressionMembers.get(nExprMemberIndex);

            // For all unhandled sub-properties
            for (int nSubPropIndex = 0; nSubPropIndex < nSubProperties; ++nSubPropIndex)
                if (!aHandledSubProperties[nSubPropIndex]) {
                    final CSSPropertyWithDefaultValue aSubProp = m_aSubProperties.get(nSubPropIndex);
                    final ICSSProperty aProperty = aSubProp.getProperty();
                    final int nMinArgs = aProperty.getMinimumArgumentCount();

                    // Always use minimum number of arguments
                    if (nExprMemberIndex + nMinArgs - 1 < nExpressionMembers) {
                        // Build sum of all members
                        final StringBuilder aSB = new StringBuilder();
                        for (int k = 0; k < nMinArgs; ++k) {
                            final String sValue = aMember.getAsCSSString(aCWS);
                            if (aSB.length() > 0)
                                aSB.append(' ');
                            aSB.append(sValue);
                        }
                        //
//                        if (aProperty.isValidValue (aSB.toString ())) {
                            // We found a match
                            final CSSExpression aExpr = new CSSExpression();
                            for (int k = 0; k < nMinArgs; ++k)
                                aExpr.addMember(aExpressionMembers.get(nExprMemberIndex + k));
                            ret.add(new CSSDeclaration(aSubProp.getProperty().getPropertyName(), aExpr));
                            nExprMemberIndex += nMinArgs - 1;

                            // Remember as handled
                            aHandledSubProperties[nSubPropIndex] = true;

                            // Next expression member
                            break;
//                        }
                    }
                }
        }

        // Assign all default values that are not present
        for (int nSubPropIndex = 0; nSubPropIndex < nSubProperties; ++nSubPropIndex)
            if (nSubPropIndex < aHandledSubProperties.length && !aHandledSubProperties[nSubPropIndex]) {
                final CSSPropertyWithDefaultValue aSubProp = m_aSubProperties.get(nSubPropIndex);
                // assign default value
                final CSSExpression aExpr = new CSSExpression();
                aExpr.addMember(new CSSExpressionMemberTermSimple(aSubProp.getDefaultValue()));
                ret.add(new CSSDeclaration(aSubProp.getProperty().getPropertyName(), aExpr));
            }

        return ret;
    }
}
