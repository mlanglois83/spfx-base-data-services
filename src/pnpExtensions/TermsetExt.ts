import { find, findIndex } from "@microsoft/sp-lodash-subset";
import { isArray } from "@pnp/common";
import { extendFactory } from "@pnp/odata";
import { defaultPath, SPRest } from "@pnp/sp";
import { spInvokableFactory, _SharePointQueryableCollection } from "@pnp/sp/sharepointqueryable";
import "@pnp/sp/taxonomy";
import { IOrderedTermInfo, ITerm, ITermInfo, ITermSet, ITermSortOrderInfo, Term, TermSet } from "@pnp/sp/taxonomy";
import { tag } from "@pnp/sp/telemetry";

// Legacy children enables retrieving reused child terms

@defaultPath("getLegacyChildren()")
// eslint-disable-next-line @typescript-eslint/class-name-casing
export class _LegacyChildren extends _SharePointQueryableCollection<ITermInfo[]> { }
// eslint-disable-next-line @typescript-eslint/no-empty-interface
export interface ILegacyChildren extends _LegacyChildren { }
export const LegacyChildren = spInvokableFactory<ILegacyChildren>(_LegacyChildren);

declare module "@pnp/sp/taxonomy" {
    interface ITermSet {
        getAllChildrenAsOrderedTreeFull: () => Promise<IOrderedTermInfo[]>;
        getLegacyChildren: () => ILegacyChildren;
    }
    interface ITerm {
        getLegacyChildren: () => ILegacyChildren;
    }
}

// child terms for term
extendFactory(Term, {
    getLegacyChildren(this: ITerm): ILegacyChildren {
        return tag.configure(LegacyChildren(this), "txt.getLegacyChildren");
    }
});

extendFactory(TermSet, {
    // child terms for termset
    getLegacyChildren(this: ITermSet): ILegacyChildren {
        return tag.configure(LegacyChildren(this), "txtg.getLegacyChildren");
    },
    // ordered terms with custom properties and custom sort order
    async getAllChildrenAsOrderedTreeFull(this: ITermSet): Promise<IOrderedTermInfo[]> {
        const setInfo = await this.select("*", "customSortOrder")();
        const tree: IOrderedTermInfo[] = [];

        const ensureOrder = (terms: IOrderedTermInfo[], sorts: ITermSortOrderInfo[], setSorts?: string[]): IOrderedTermInfo[] => {

            // handle no custom sort information present
            if (!isArray(sorts) && !isArray(setSorts)) {
                return terms;
            }

            let ordering: string[] = null;
            if (sorts === null && setSorts.length > 0) {
                ordering = [...setSorts];
            } else {
                const index = findIndex(sorts, v => v.setId === setInfo.id);
                if (index >= 0) {
                    ordering = [...sorts[index].order];
                }
            }

            if (ordering !== null) {
                const orderedChildren = [];
                ordering.forEach(o => {
                    const found = find(terms, ch => o === ch.id);
                    if (found) {
                        orderedChildren.push(found);
                    }
                });
                // we have a case where if a set is ordered and a term is added to that set
                // AND the ordering information hasn't been updated the new term will not have
                // any associated ordering information. See #1547 which reported this. So here we
                // append any terms remaining in "terms" not in "orderedChildren" to the end of "orderedChildren"
                orderedChildren.push(...terms.filter(info => ordering.indexOf(info.id) < 0));

                return orderedChildren;
            }
            return terms;
        };

        const visitor = async (source: ITermSet | ITerm, parent: IOrderedTermInfo[]): Promise<void> => {

            const children = await source.getLegacyChildren().select("*", "customSortOrder","properties", "localProperties").get();

            for (let i = 0; i < children.length; i++) {

                const child = children[i];

                const orderedTerm = {
                    children: new Array<IOrderedTermInfo>(),
                    defaultLabel: find(child.labels,l => l.isDefault).name,
                    ...child,
                };

                if (child.childrenCount > 0) {
                    await visitor(this.getTermById(children[i].id), orderedTerm.children);
                    orderedTerm.children = ensureOrder(orderedTerm.children, child.customSortOrder);
                }

                parent.push(orderedTerm);
            }
        };

        await visitor(this, tree);

        return ensureOrder(tree, null, setInfo.customSortOrder);
    }
});

export const sp = new SPRest();